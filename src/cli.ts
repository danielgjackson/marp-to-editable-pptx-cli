import { unlinkSync } from 'node:fs'
// Reused directly from the upstream VS Code extension's sources.
import { detectBrowserPath } from 'marp-to-editable-pptx/src/native-pptx/browser'
import { convertFile, pendingTempFiles } from './convert'
import { CLIError, CLIErrorCode } from './error'
import { parseCliArgs, USAGE, UsageError, type CliOptions } from './options'
import { planJobs, STDIO, type Job } from './plan'

declare const __CLI_VERSION__: string
declare const __UPSTREAM_VERSION__: string

function removePendingTempFiles(): void {
  for (const file of pendingTempFiles) {
    try {
      unlinkSync(file)
    } catch {
      // already gone
    }
  }
  pendingTempFiles.clear()
}

const label = (p: string, stdio: string) => (p === STDIO ? stdio : p)

function describe(job: Job): string {
  return `${label(job.input, '<stdin>')} -> ${label(job.output, '<stdout>')}`
}

async function run(opts: CliOptions): Promise<number> {
  const info = (msg: string) => {
    if (!opts.quiet) process.stderr.write(msg + '\n')
  }
  const error = (msg: string) => process.stderr.write(`error: ${msg}\n`)

  const jobs = planJobs(opts)
  if (jobs.length === 0) return 0

  if (!opts.dryRun && process.stdout.isTTY && jobs.some((j) => j.output === STDIO)) {
    throw new UsageError(
      'refusing to write binary .pptx data to a terminal; redirect standard output or use -o FILE',
    )
  }

  if (opts.dryRun) {
    const verb: Record<Job['action'], string> = {
      create: 'convert',
      overwrite: 'convert (overwrite)',
      'up-to-date': 'skip (up to date)',
      exists: 'skip (output exists; use --force or --update)',
    }
    for (const job of jobs) info(`${verb[job.action]}: ${describe(job)}`)
    return 0
  }

  const todo = jobs.filter((j) => j.action === 'create' || j.action === 'overwrite')
  let browserPath: string | undefined
  if (todo.length > 0) {
    browserPath = detectBrowserPath(opts.browser, opts.browserPath)
    if (!browserPath) {
      error(
        opts.browserPath
          ? `browser executable not found: ${opts.browserPath}`
          : 'could not find a Chromium-based browser (Chrome or Edge), which is required for PPTX export. ' +
              'Install one, or pass --browser-path (or set CHROME_PATH).',
      )
      return 1
    }
  }

  // Ctrl+C: abort the current conversion. Temporary files are removed
  // synchronously on exit, because puppeteer's own SIGINT handler (enabled by
  // upstream's launch options) calls process.exit() straight away.
  // The listeners are removed again afterwards, so that marpCli() leaves a
  // host application's process as it found it.
  const controller = new AbortController()
  const onSignal = (sig: NodeJS.Signals) => {
    if (controller.signal.aborted) process.exit(130)
    info(`\nReceived ${sig}, stopping...`)
    controller.abort(new DOMException('Conversion cancelled', 'AbortError'))
  }
  process.on('exit', removePendingTempFiles)
  process.on('SIGINT', onSignal)
  process.on('SIGTERM', onSignal)
  try {
    return await runJobs(jobs, opts, browserPath, controller, info, error)
  } finally {
    process.off('SIGINT', onSignal)
    process.off('SIGTERM', onSignal)
    process.off('exit', removePendingTempFiles)
    removePendingTempFiles()
  }
}

async function runJobs(
  jobs: Job[],
  opts: CliOptions,
  browserPath: string | undefined,
  controller: AbortController,
  info: (msg: string) => void,
  error: (msg: string) => void,
): Promise<number> {

  let converted = 0
  let skipped = 0
  let failed = 0
  let stopped = false // set after the first error, unless --keep-going
  const multi = jobs.length > 1
  const fail = (msg: string) => {
    failed++
    error(msg)
    if (!opts.keepGoing) stopped = true
  }

  const processJob = async (job: Job) => {
    if (job.action === 'up-to-date') {
      skipped++
      if (opts.verbose) info(`Up to date: ${describe(job)}`)
      return
    }
    if (job.action === 'exists') {
      fail(`output exists, not overwriting (use --force or --update): ${job.output}`)
      return
    }

    const started = Date.now()
    info(`Converting: ${describe(job)}`)
    if (opts.verbose && job.config) info(`  using Marp config ${job.config}`)
    try {
      await convertFile({
        input: job.input,
        output: job.output,
        config: job.config,
        browserPath: browserPath!,
        opts,
        signal: controller.signal,
        log: (msg) => {
          if (opts.verbose) process.stderr.write(`  ${label(job.input, '<stdin>')}: ${msg}\n`)
        },
      })
      converted++
      if (opts.verbose) {
        info(`  ${label(job.input, '<stdin>')}: done in ${((Date.now() - started) / 1000).toFixed(1)}s`)
      }
    } catch (err) {
      if (controller.signal.aborted) return
      fail(`${describe(job)}: ${err instanceof Error ? err.message : String(err)}`)
    }
  }

  // Up to --parallel conversions at a time, started in input order. Each one
  // launches its own headless browser.
  let next = 0
  const worker = async () => {
    while (next < jobs.length && !stopped && !controller.signal.aborted) {
      await processJob(jobs[next++])
    }
  }
  await Promise.all(Array.from({ length: Math.min(opts.jobs, jobs.length) }, worker))

  if (controller.signal.aborted) {
    error('cancelled')
    return 130
  }

  if (multi || opts.verbose) {
    const parts = [`${converted} converted`]
    if (skipped) parts.push(`${skipped} up to date`)
    if (failed) parts.push(`${failed} failed`)
    const notRun = jobs.length - converted - skipped - failed
    if (notRun) parts.push(`${notRun} not attempted`)
    info(parts.join(', '))
  }
  return failed > 0 ? 1 : 0
}

/** Options for {@link marpCli}. */
export interface MarpCLIAPIOptions {
  /**
   * Read Markdown from standard input when no input is given, as the
   * command-line tool does when data is piped to it. Default: `false`, so a
   * call without inputs fails with a usage error instead of waiting on the
   * host process's stdin. An explicit `-` input always reads stdin.
   */
  stdin?: boolean
}

/**
 * Run the CLI with the given arguments (without the leading `node` and script
 * path), as Marp CLI's `marpCli()` does. Messages go to stderr and stdout as
 * they would from the command line.
 *
 * Resolves with the exit status the command would have returned: 0 on
 * success, 1 if a conversion failed or an existing output was not
 * overwritten, 130 if interrupted. Rejects with a {@link CLIError} (errorCode
 * {@link CLIErrorCode.INVALID_OPTIONS}) for invalid usage, and with other
 * errors for unexpected failures.
 */
export async function marpCli(argv: string[], options: MarpCLIAPIOptions = {}): Promise<number> {
  const opts = parseCliArgs(argv)

  if (opts.help) {
    process.stdout.write(USAGE)
    return 0
  }
  if (opts.version) {
    process.stdout.write(
      `marp-to-editable-pptx-cli ${__CLI_VERSION__} (marp-to-editable-pptx ${__UPSTREAM_VERSION__})\n`,
    )
    return 0
  }
  if (opts.inputs.length === 0 && (!options.stdin || process.stdin.isTTY)) {
    throw new UsageError('no input files given')
  }

  return run(opts)
}

async function main(): Promise<void> {
  const argv = process.argv.slice(2)
  try {
    process.exitCode = await marpCli(argv, { stdin: true })
  } catch (err) {
    if (!(err instanceof CLIError)) throw err
    if (err.errorCode === CLIErrorCode.INVALID_OPTIONS) {
      // Without any input or options, show the usage instead.
      if (argv.length === 0) process.stderr.write(USAGE)
      else process.stderr.write(`error: ${err.message}\nTry --help for usage.\n`)
    } else {
      process.stderr.write(`error: ${err.message}\n`)
    }
    process.exitCode = err.errorCode
  }
}

// Run only when executed as a program, so that the bundle can also be
// require()d / imported as a library.
if (require.main === module) {
  main().catch((err) => {
    process.stderr.write(`error: ${err instanceof Error ? err.stack : String(err)}\n`)
    process.exit(1)
  })
}

export { CLIError, CLIErrorCode }
// Used by the tests to find a browser for the end-to-end test.
export { detectBrowserPath }
