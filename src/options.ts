import { parseArgs } from 'node:util'
import { CLIError, CLIErrorCode } from './error'

export type BrowserPreference = 'auto' | 'chrome' | 'edge'

export interface CliOptions {
  inputs: string[]
  output?: string
  outputDir?: string
  force: boolean
  update: boolean
  dryRun: boolean
  recursive: boolean
  keepGoing: boolean
  /** Maximum parallel conversions (-P, --parallel). */
  jobs: number
  config?: string
  noConfig: boolean
  theme?: string
  themeSet: string[]
  html?: boolean
  allowLocalFiles: boolean
  browser: BrowserPreference
  browserPath?: string
  marpArgs: string[]
  quiet: boolean
  /** -d, --debug (Marp CLI) or --verbose (pandoc). */
  verbose: boolean
  help: boolean
  version: boolean
}

/** Invalid usage (exit status 2). */
export class UsageError extends CLIError {
  constructor(message: string) {
    super(message, CLIErrorCode.INVALID_OPTIONS)
  }
}

/** Marp CLI's default for --parallel. */
export const DEFAULT_PARALLEL = 5

export const USAGE = `Usage:
  marp-to-editable-pptx-cli [options] <files...> [-- marp-cli-args...]
  marp-to-editable-pptx-cli [options] -I <dir>

Convert Marp Markdown into editable PowerPoint (.pptx) files. Options follow
Marp CLI ("marp --pptx --pptx-editable ..."), so the same command lines work;
PPTX output is always editable, and --pptx / --pptx-editable are the defaults.

A file of "-", or no file with piped input, reads Markdown from standard input.
Each slides.md is written beside it as slides.pptx unless -o is given. Unlike
Marp CLI, existing files are never overwritten without --force or --update.

Basic options (as in Marp CLI):
  -o, --output <path>      Output file ("-" for standard output), or the
                           output directory when --input-dir is given
  -I, --input-dir <dir>    Convert every Markdown file below <dir>, keeping
                           the folder structure in the output directory
  -c, --config-file, --config <file>
                           Marp config file (default: look for .marprc.yml,
                           marp.config.js, ... in each deck's folder, then its
                           parents up to the project root)
      --no-config-file, --no-config
                           Do not look for a Marp config file
  -P, --parallel [num]     Convert up to <num> files at once (default: ${DEFAULT_PARALLEL})
      --no-parallel        Convert one file at a time
  -d, --debug              Show Marp CLI output and conversion progress
  -v, --version            Show versions
  -h, --help               Show help

Converter, browser and Marp options (as in Marp CLI):
      --pptx               Convert to PowerPoint (always on)
      --pptx-editable      Generate editable PPTX (always on)
      --allow-local-files  Allow access to local files from Markdown
      --html, --no-html    Enable or disable HTML tags in Markdown
      --theme <name|file>  Override the theme by name or CSS file
      --theme-set <path>   Additional theme CSS file or directory (repeatable)
      --browser <kind>     auto, chrome or edge (default: auto)
      --browser-path <path>
                           Browser executable (default: $CHROME_PATH, or found
                           automatically)
  -- <args...>             Pass any remaining arguments to Marp CLI's
                           Markdown-to-HTML step

Batch and overwrite options (extensions):
  -f, --force              Overwrite existing output files
  -u, --update             Only convert when the output is missing or older
                           than its input (or its config/theme files); stale
                           outputs are overwritten
  -n, --dry-run            Show what would be converted, without converting
  -k, --keep-going         Continue after an error (default: stop at the first)
  -r, --recursive          Include subfolders of directories given as <files>
      --output-dir <dir>   Write outputs for <files> into <dir>
  -q, --quiet              Only print errors
      --verbose            Same as --debug

Examples:
  marp-to-editable-pptx-cli --pptx --pptx-editable slides.md
  marp-to-editable-pptx-cli slides.md -o deck.pptx
  marp-to-editable-pptx-cli -u -I decks/ -o build/
  cat slides.md | marp-to-editable-pptx-cli -o - > deck.pptx
`

/** Marp CLI outputs that are not editable PPTX, and so are not supported. */
const UNSUPPORTED_OUTPUTS = ['pdf', 'notes', 'image', 'images'] as const
const UNSUPPORTED_MODES = ['watch', 'server', 'preview'] as const

/**
 * `-P` / `--parallel` take an optional number, as in Marp CLI. parseArgs
 * needs a value, so a bare one is given the default here.
 */
function normalizeParallel(argv: string[]): string[] {
  const out: string[] = []
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]
    if ((arg === '-P' || arg === '--parallel') && !/^\d+$/.test(argv[i + 1] ?? '')) {
      out.push(`--parallel=${DEFAULT_PARALLEL}`)
    } else {
      out.push(arg)
    }
  }
  return out
}

export function parseCliArgs(argv: string[]): CliOptions {
  // Everything after a bare `--` is passed through to Marp CLI untouched.
  const sep = argv.indexOf('--')
  const ownArgs = normalizeParallel(sep >= 0 ? argv.slice(0, sep) : argv)
  const marpArgs = sep >= 0 ? argv.slice(sep + 1) : []

  let parsed
  try {
    parsed = parseArgs({
      args: ownArgs,
      allowPositionals: true,
      allowNegative: true,
      strict: true,
      options: {
        // Marp CLI basic options
        output: { type: 'string', short: 'o' },
        'input-dir': { type: 'string', short: 'I' },
        'config-file': { type: 'string', short: 'c' },
        config: { type: 'string' },
        // Declared so the negated forms are accepted: allowNegative alone
        // only applies to boolean options. They are reported as `false` on
        // the positive name (e.g. `config: false`).
        'no-config-file': { type: 'boolean' },
        'no-config': { type: 'boolean' },
        parallel: { type: 'string', short: 'P' },
        'no-parallel': { type: 'boolean' },
        debug: { type: 'boolean', short: 'd' },
        version: { type: 'boolean', short: 'v' },
        help: { type: 'boolean', short: 'h' },
        // Marp CLI converter / Marp / browser options
        pptx: { type: 'boolean' },
        'pptx-editable': { type: 'boolean' },
        'allow-local-files': { type: 'boolean' },
        html: { type: 'boolean' },
        theme: { type: 'string' },
        'theme-set': { type: 'string', multiple: true },
        browser: { type: 'string' },
        'browser-path': { type: 'string' },
        // Marp CLI options that are recognised only to be rejected clearly
        pdf: { type: 'boolean' },
        notes: { type: 'boolean' },
        image: { type: 'string' },
        images: { type: 'string' },
        watch: { type: 'boolean', short: 'w' },
        server: { type: 'boolean', short: 's' },
        preview: { type: 'boolean', short: 'p' },
        // Extensions
        force: { type: 'boolean', short: 'f' },
        update: { type: 'boolean', short: 'u' },
        'dry-run': { type: 'boolean', short: 'n' },
        'keep-going': { type: 'boolean', short: 'k' },
        recursive: { type: 'boolean', short: 'r' },
        'output-dir': { type: 'string' },
        quiet: { type: 'boolean', short: 'q' },
        verbose: { type: 'boolean' },
      },
    })
  } catch (error) {
    throw new UsageError(error instanceof Error ? error.message : String(error))
  }

  const { values: v, positionals } = parsed

  for (const name of UNSUPPORTED_OUTPUTS) {
    if (v[name] !== undefined && v[name] !== false) {
      throw new UsageError(
        `--${name} is not supported: this tool only produces editable PPTX (use Marp CLI for other formats)`,
      )
    }
  }
  if (v.pptx === false || v['pptx-editable'] === false) {
    throw new UsageError(
      `--no-${v.pptx === false ? 'pptx' : 'pptx-editable'} is not supported: this tool only produces editable PPTX (use Marp CLI for image-based PPTX)`,
    )
  }
  for (const name of UNSUPPORTED_MODES) {
    if (v[name]) throw new UsageError(`--${name} is not supported by this tool`)
  }

  // --config / --config-file are aliases; with allowNegative, the --no-*
  // forms are reported as `false`.
  const configValues = [v['config-file'], v.config] as (string | boolean | undefined)[]
  const configs = configValues.filter((c): c is string => typeof c === 'string')
  const noConfig = configValues.includes(false)
  if (configs.length > 1) throw new UsageError('give the config file only once')
  if (configs.length > 0 && noConfig) {
    throw new UsageError('--config-file and --no-config-file cannot be used together')
  }

  const parallelValue = v.parallel as string | boolean | undefined
  let jobs = DEFAULT_PARALLEL
  if (parallelValue === false) jobs = 1
  else if (typeof parallelValue === 'string') {
    jobs = Number(parallelValue)
    if (!Number.isInteger(jobs) || jobs < 0) {
      throw new UsageError(`--parallel must be a whole number (got "${parallelValue}")`)
    }
    jobs = Math.max(1, jobs) // Marp CLI treats 0 as 1
  }

  const browser = (v.browser ?? 'auto') as string
  if (!['auto', 'chrome', 'edge'].includes(browser)) {
    throw new UsageError(
      `--browser must be one of auto, chrome, edge (got "${browser}")`,
    )
  }

  // -I <dir>: Marp CLI's batch mode. All Markdown below <dir> is converted,
  // and -o (if given) names the output directory.
  const inputDir = v['input-dir']
  let inputs = positionals
  let output = v.output
  let outputDir = v['output-dir']
  if (inputDir !== undefined) {
    if (positionals.length > 0) {
      throw new UsageError('give either --input-dir or input files, not both')
    }
    if (output === '-') {
      throw new UsageError('--input-dir cannot write to standard output')
    }
    if (output !== undefined && outputDir !== undefined) {
      throw new UsageError('--output and --output-dir cannot be used together')
    }
    inputs = [inputDir]
    outputDir = output ?? outputDir
    output = undefined
  } else if (output !== undefined && outputDir !== undefined) {
    throw new UsageError('--output and --output-dir cannot be used together')
  }

  const opts: CliOptions = {
    inputs,
    output,
    outputDir,
    force: !!v.force,
    update: !!v.update,
    dryRun: !!v['dry-run'],
    recursive: !!v.recursive || inputDir !== undefined,
    keepGoing: !!v['keep-going'],
    jobs,
    config: configs[0],
    noConfig,
    theme: v.theme,
    themeSet: v['theme-set'] ?? [],
    html: v.html,
    allowLocalFiles: !!v['allow-local-files'],
    browser: browser as BrowserPreference,
    browserPath: v['browser-path'] ?? (process.env.CHROME_PATH || undefined),
    marpArgs,
    quiet: !!v.quiet,
    verbose: !!v.debug || !!v.verbose,
    help: !!v.help,
    version: !!v.version,
  }

  if (opts.quiet && opts.verbose) {
    throw new UsageError('--quiet and --debug/--verbose cannot be used together')
  }
  return opts
}
