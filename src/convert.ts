import { randomBytes } from 'node:crypto'
import { mkdir, rename, unlink, writeFile } from 'node:fs/promises'
import path from 'node:path'
import { format } from 'node:util'
import { marpCli } from '@marp-team/marp-cli'
// Reused directly from the upstream VS Code extension's sources.
import { generateNativePptx } from 'marp-to-editable-pptx/src/native-pptx/index'
import { optimizePptx } from './optimize'
import type { CliOptions } from './options'
import { STDIO } from './plan'

export interface ConvertRequest {
  /** Markdown file path, or "-" for standard input. */
  input: string
  /** PPTX file path, or "-" for standard output. */
  output: string
  /** Marp config file to use (from --config or found next to the input). */
  config?: string
  browserPath: string
  opts: CliOptions
  signal: AbortSignal
  log: (message: string) => void
}

/** Temporary files that must be removed if the process is interrupted. */
export const pendingTempFiles = new Set<string>()

function tempName(dir: string, suffix: string): string {
  const id = randomBytes(6).toString('hex')
  const file = path.join(dir, `.marp-editable-pptx-${id}${suffix}`)
  pendingTempFiles.add(file)
  return file
}

async function removeTemp(file: string): Promise<void> {
  pendingTempFiles.delete(file)
  try {
    await unlink(file)
  } catch {
    // already gone
  }
}

async function readStdin(): Promise<Buffer> {
  const chunks: Buffer[] = []
  for await (const chunk of process.stdin) chunks.push(chunk as Buffer)
  return Buffer.concat(chunks)
}

/** Build the Marp CLI argument list for the Markdown → HTML step. */
export function marpCliArgs(
  input: string,
  htmlPath: string,
  config: string | undefined,
  opts: CliOptions,
): string[] {
  const args = [input, '-o', htmlPath]
  // Config lookup is done by us (per input), so stop Marp CLI from also
  // looking in the current directory when none applies.
  if (config !== undefined) args.push('--config-file', config)
  else args.push('--no-config-file')
  if (opts.theme !== undefined) args.push('--theme', opts.theme)
  for (const theme of opts.themeSet) args.push('--theme-set', theme)
  if (opts.html === true) args.push('--html')
  if (opts.html === false) args.push('--no-html')
  if (opts.allowLocalFiles) args.push('--allow-local-files')
  args.push(...opts.marpArgs)
  return args
}

/**
 * Marp CLI runs in this process and is captured by patching console.*, so
 * runs are serialized: with --parallel, only the (slow) browser step is parallel.
 */
let marpQueue: Promise<unknown> = Promise.resolve()
function serialized<T>(fn: () => Promise<T>): Promise<T> {
  const result = marpQueue.then(fn, fn)
  marpQueue = result.catch(() => {})
  return result
}

/**
 * Run Marp CLI while capturing its console output (it logs via console.*),
 * so that it can be shown only in verbose mode or on failure.
 */
async function runMarpCli(
  args: string[],
): Promise<{ exitCode: number; output: string }> {
  const captured: string[] = []
  const methods = ['log', 'info', 'warn', 'error'] as const
  const originals = methods.map((m) => console[m])
  for (const m of methods) {
    console[m] = (...data: unknown[]) => {
      captured.push(format(...data))
    }
  }
  try {
    const exitCode = await marpCli(args, {})
    return { exitCode, output: captured.join('\n') }
  } finally {
    methods.forEach((m, i) => (console[m] = originals[i]))
  }
}

/** Convert a single Marp Markdown file to an editable PPTX file. */
export async function convertFile(req: ConvertRequest): Promise<void> {
  const { opts, signal, log } = req
  const cleanup: string[] = []

  try {
    // Marp CLI resolves relative image paths from the HTML output location,
    // so (like the VS Code extension) put the temporary HTML next to the
    // source Markdown. Standard input is treated as living in the cwd.
    let inputPath = req.input
    if (inputPath === STDIO) {
      inputPath = tempName(process.cwd(), '.md')
      cleanup.push(inputPath)
      await writeFile(inputPath, await readStdin())
    }
    const htmlPath = tempName(path.dirname(path.resolve(inputPath)), '.html')
    cleanup.push(htmlPath)

    // Step 1: Markdown → HTML (Marp CLI)
    log('Rendering Markdown with Marp CLI')
    const { exitCode, output } = await serialized(() =>
      runMarpCli(marpCliArgs(inputPath, htmlPath, req.config, opts)),
    )
    if (opts.verbose && output) process.stderr.write(output + '\n')
    if (exitCode !== 0) {
      throw new Error(
        `Marp CLI exited with code ${exitCode}` +
          (output && !opts.verbose ? `:\n${output}` : ''),
      )
    }
    signal.throwIfAborted()

    // Step 2: HTML → editable PPTX (marp-to-editable-pptx)
    let pptx = await generateNativePptx({
      htmlPath,
      browserPath: req.browserPath,
      signal,
      onProgress: log,
    })
    signal.throwIfAborted()

    // Step 3: shrink the package (shared media, compression); see optimize.ts.
    if (opts.optimize) {
      let result
      try {
        result = await optimizePptx(pptx)
      } catch (error) {
        throw new Error(
          `could not optimize the PPTX (use --no-optimize to skip this step): ${
            error instanceof Error ? error.message : String(error)
          }`,
        )
      }
      const kb = (n: number) => `${Math.round(n / 1024)} KB`
      log(
        `Optimized PPTX: ${kb(pptx.length)} -> ${kb(result.data.length)}` +
          ` (media files: ${result.mediaBefore} -> ${result.mediaAfter})`,
      )
      pptx = result.data
      signal.throwIfAborted()
    }

    // Step 4: write the result (atomically, via a temporary file + rename).
    if (req.output === STDIO) {
      await new Promise<void>((resolve, reject) =>
        process.stdout.write(pptx, (err) => (err ? reject(err) : resolve())),
      )
      return
    }
    const outDir = path.dirname(path.resolve(req.output))
    await mkdir(outDir, { recursive: true })
    const tmpOut = tempName(outDir, '.pptx.tmp')
    cleanup.push(tmpOut)
    await writeFile(tmpOut, pptx)
    signal.throwIfAborted()
    await rename(tmpOut, req.output)
  } finally {
    await Promise.all(cleanup.map(removeTemp))
  }
}
