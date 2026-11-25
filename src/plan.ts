import { readdirSync, statSync, type Stats } from 'node:fs'
import path from 'node:path'
import { cosmiconfigSync } from 'cosmiconfig'
import { UsageError, type CliOptions } from './options'

export const STDIO = '-'

const MARKDOWN_EXTENSIONS = new Set(['.md', '.markdown'])
const SKIPPED_DIRECTORIES = new Set(['node_modules'])

export type JobAction =
  /** Output does not exist (or is stdout): convert. */
  | 'create'
  /** Output exists and will be replaced (--force, or stale with --update). */
  | 'overwrite'
  /** --update: output is at least as new as its inputs. */
  | 'up-to-date'
  /** Output exists and neither --force nor --update was given. */
  | 'exists'

export interface Job {
  /** Markdown path, or "-" for standard input. */
  input: string
  /** PPTX path, or "-" for standard output. */
  output: string
  action: JobAction
  /** Marp config file for this input (explicit --config or discovered). */
  config?: string
}

interface SourceFile {
  file: string
  /** Directory that `file` is resolved relative to when using --output-dir. */
  base: string
}

function tryStat(p: string): Stats | undefined {
  try {
    return statSync(p)
  } catch {
    return undefined
  }
}

export function isMarkdownFile(name: string): boolean {
  return MARKDOWN_EXTENSIONS.has(path.extname(name).toLowerCase())
}

/** List Markdown files in `dir` (sorted), optionally recursing. */
export function findMarkdownFiles(dir: string, recursive: boolean): string[] {
  const found: string[] = []
  const entries = readdirSync(dir, { withFileTypes: true }).sort((a, b) =>
    a.name.localeCompare(b.name),
  )
  for (const entry of entries) {
    if (entry.name.startsWith('.')) continue
    const full = path.join(dir, entry.name)
    if (entry.isDirectory()) {
      if (recursive && !SKIPPED_DIRECTORIES.has(entry.name)) {
        found.push(...findMarkdownFiles(full, recursive))
      }
    } else if (isMarkdownFile(entry.name) && tryStat(full)?.isFile()) {
      found.push(full)
    }
  }
  return found
}

export function replaceExtension(file: string, ext: string): string {
  const parsed = path.parse(file)
  return path.join(parsed.dir, parsed.name + ext)
}

/**
 * Find the Marp config that applies to a deck, looking in its folder and then
 * each parent up to the project root (the nearest folder with package.json).
 * Upstream's VS Code extension also looks from the deck's folder; Marp CLI on
 * its own only checks the current directory.
 */
export function makeConfigFinder(): (dir: string) => string | undefined {
  const explorer = cosmiconfigSync('marp', { searchStrategy: 'project' })
  const cache = new Map<string, string | undefined>()
  return (dir) => {
    const key = path.resolve(dir)
    if (!cache.has(key)) {
      const found = explorer.search(key)
      cache.set(key, found && !found.isEmpty ? found.filepath : undefined)
    }
    return cache.get(key)
  }
}

/**
 * Latest modification time (ms) of the config and theme files that affect
 * every conversion, so that --update also reconverts when they change.
 */
export function sharedDependencyMtime(opts: CliOptions): number {
  const paths = [opts.theme, ...opts.themeSet].filter(
    (p): p is string => !!p,
  )
  let latest = 0
  for (const p of paths) {
    const st = tryStat(p)
    if (!st) continue
    latest = Math.max(latest, st.mtimeMs)
    if (st.isDirectory()) {
      for (const name of readdirSync(p)) {
        if (path.extname(name).toLowerCase() !== '.css') continue
        latest = Math.max(latest, tryStat(path.join(p, name))?.mtimeMs ?? 0)
      }
    }
  }
  return latest
}

/** Expand the command-line inputs into concrete conversion jobs. */
export function planJobs(opts: CliOptions): Job[] {
  const inputs = opts.inputs.length > 0 ? opts.inputs : [STDIO]

  // 1. Expand inputs into source files.
  const sources: SourceFile[] = []
  const seen = new Set<string>()
  let stdinCount = 0
  for (const input of inputs) {
    if (input === STDIO) {
      if (++stdinCount > 1) {
        throw new UsageError('standard input ("-") can only be given once')
      }
      sources.push({ file: STDIO, base: '.' })
      continue
    }
    const st = tryStat(input)
    if (!st) throw new UsageError(`input not found: ${input}`)
    const add = (file: string, base: string) => {
      const key = path.resolve(file)
      if (seen.has(key)) return
      seen.add(key)
      sources.push({ file, base })
    }
    if (st.isDirectory()) {
      const files = findMarkdownFiles(input, opts.recursive)
      if (files.length === 0 && !opts.quiet) {
        process.stderr.write(`warning: no Markdown files found in ${input}\n`)
      }
      for (const file of files) add(file, input)
    } else {
      add(input, path.dirname(input))
    }
  }

  // 2. Map each source to an output path.
  if (opts.output !== undefined && sources.length !== 1) {
    throw new UsageError(
      `--output can only be used with a single input (got ${sources.length}); use -I <dir> -o <dir>, or --output-dir, for multiple inputs`,
    )
  }
  const jobs: Job[] = sources.map(({ file, base }) => {
    let output: string
    if (opts.output !== undefined) {
      output = opts.output
    } else if (file === STDIO) {
      throw new UsageError(
        'reading from standard input requires --output (use "-o -" to write to standard output)',
      )
    } else if (opts.outputDir !== undefined) {
      output = path.join(
        opts.outputDir,
        replaceExtension(path.relative(base, file), '.pptx'),
      )
    } else {
      output = replaceExtension(file, '.pptx')
    }
    return { input: file, output, action: 'create' as JobAction }
  })

  // 3. Sanity checks: no output may clobber an input or another output.
  const outputs = new Map<string, string>()
  for (const job of jobs) {
    if (job.output === STDIO) continue
    const key = path.resolve(job.output)
    if (job.input !== STDIO && key === path.resolve(job.input)) {
      throw new UsageError(`output would overwrite its input: ${job.input}`)
    }
    const other = outputs.get(key)
    if (other !== undefined) {
      throw new UsageError(
        `${other} and ${job.input} would both be written to ${job.output}`,
      )
    }
    outputs.set(key, job.input)
  }

  // 4. Work out which Marp config file applies to each input.
  if (opts.config !== undefined) {
    if (!tryStat(opts.config)?.isFile()) {
      throw new UsageError(`config file not found: ${opts.config}`)
    }
    for (const job of jobs) job.config = opts.config
  } else if (!opts.noConfig) {
    const findConfig = makeConfigFinder()
    for (const job of jobs) {
      const dir = job.input === STDIO ? process.cwd() : path.dirname(job.input)
      job.config = findConfig(dir)
    }
  }

  // 5. Decide what to do with each job based on existing outputs.
  const depMtime = opts.update ? sharedDependencyMtime(opts) : 0
  for (const job of jobs) {
    if (job.output === STDIO) continue
    const outStat = tryStat(job.output)
    if (!outStat) continue
    if (outStat.isDirectory()) {
      throw new UsageError(`output path is a directory: ${job.output}`)
    }
    if (opts.update && job.input !== STDIO) {
      const inMtime = statSync(job.input).mtimeMs
      const configMtime = job.config ? (tryStat(job.config)?.mtimeMs ?? 0) : 0
      job.action =
        outStat.mtimeMs >= Math.max(inMtime, depMtime, configMtime)
          ? 'up-to-date'
          : 'overwrite'
    } else if (opts.force || opts.update) {
      // (--update with standard input: nothing to compare, so always convert)
      job.action = 'overwrite'
    } else {
      job.action = 'exists'
    }
  }

  return jobs
}
