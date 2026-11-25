**(The code in this repository has been written by Claude AI)**

# marp-to-editable-pptx-cli

A command-line tool that converts [Marp](https://marp.app/) Markdown slide decks into **editable** *PowerPoint* (`.pptx`) files. In the output, text, lists and tables are real *PowerPoint* objects, not screenshots -- unlike the standard [marp-cli](https://github.com/marp-team/marp-cli) `--pptx` output.

The experimental `--pptx-editable` option in [marp-cli](https://github.com/marp-team/marp-cli) requires *LibreOffice*; while this tool is a thin CLI wrapper around the conversion engine of [KatsuYuzu/marp-to-editable-pptx](https://github.com/KatsuYuzu/marp-to-editable-pptx) (a VS Code extension), which instead uses a headless browser.  This tool exposes the same conversion from a terminal, with a [Marp CLI](https://github.com/marp-team/marp-cli)-compatible interface and support for batch conversion.


## Requirements

- **Node.js 20.16** or newer.
- A **Chromium-based browser** (Google Chrome, Microsoft Edge or Chromium). It is used headlessly to lay out the slides, and it is found automatically in the standard install locations. Otherwise, pass `--browser-path` or set `CHROME_PATH`.

## Quick start

Run it straight from GitHub with `npx`, with nothing to install:

```sh
npx github:danielgjackson/marp-to-editable-pptx-cli slides.md
```

This writes `slides.pptx` next to `slides.md`. The full repository URL works too:

```sh
npx git+https://github.com/danielgjackson/marp-to-editable-pptx-cli.git slides.md
```

The first run takes a little longer because npm clones the tool and installs its dependencies. Later runs use npm's cache.

To install the commands permanently:

```sh
npm install --global github:danielgjackson/marp-to-editable-pptx-cli
marp-to-editable-pptx-cli slides.md      # full name
marp2pptx slides.md                      # short alias
```

## Usage

```text
marp-to-editable-pptx-cli [options] <files...> [-- marp-cli-args...]
marp-to-editable-pptx-cli [options] -I <dir>
```

The options follow [Marp CLI](https://github.com/marp-team/marp-cli), so an existing Marp command line works by swapping the command name:

```sh
marp                      --pptx --pptx-editable slide-deck.md   # Marp CLI (needs LibreOffice)
marp-to-editable-pptx-cli --pptx --pptx-editable slide-deck.md   # this tool
```

The output is always editable PPTX, so `--pptx` and `--pptx-editable` are optional. `marp2pptx slide-deck.md` does the same thing. Only the Marp CLI options that matter for this conversion are supported, and the ones for other outputs (`--pdf`, `--image`, `--notes`, `--no-pptx-editable`, `--watch`, `--server`, …) are rejected with an error. Use Marp CLI itself for those.

Differences from Marp CLI:

- **Existing files are not overwritten** unless you add `--force` or `--update` (see below). Marp CLI overwrites them.
- Each deck's Marp config file is looked up from the deck's own folder (see [Marp configuration](#marp-configuration)).
- A few batch options are added: `-f`, `-u`, `-n`, `-k`, `-r`, `-q` and `--output-dir`. They don't clash with any Marp CLI option.

Inputs can be:

- Markdown files (any extension). Each `slides.md` is written beside it as `slides.pptx`, unless `-o` names the output.
- A **directory**, with `-I <dir>` as in Marp CLI: every `.md` / `.markdown` file below it is converted, keeping the folder structure. `-o <dir>` then names the output directory; without it, each output goes beside its input.
- A directory given as a plain argument: the Markdown files directly inside it are converted, and `-r` includes subfolders.
- `-`, meaning standard input. With no input and piped data, Markdown is also read from standard input. `-o -` writes the PPTX to standard output.

Hidden folders (names starting with `.`) and `node_modules` are skipped.

### Options

Options shared with Marp CLI:

| Option | Description |
| --- | --- |
| `-o, --output <path>` | Output file (`-` = standard output), or the output **directory** when `-I` is used. With several input files, use `-I` or `--output-dir` instead. |
| `-I, --input-dir <dir>` | Convert every Markdown file below `<dir>`, keeping the folder structure. |
| `-c, --config-file, --config <file>` | Use this [Marp CLI config file](https://github.com/marp-team/marp-cli#configuration-file) for every input. |
| `--no-config-file, --no-config` | Don't look for a Marp config file. |
| `-P, --parallel [num]` | Convert up to `num` files at once. The default is `5`, as in Marp CLI. Each conversion runs its own headless browser. |
| `--no-parallel` | Convert one file at a time. |
| `-d, --debug` | Show Marp CLI output, the config file used, and progress for each step. |
| `--pptx`, `--pptx-editable` | Accepted for compatibility. They are always on. |
| `--allow-local-files` | Allow the Markdown to access local files. |
| `--html`, `--no-html` | Enable or disable HTML tags in the Markdown. |
| `--theme <name\|file>` | Override the theme, by name or by CSS file. |
| `--theme-set <path>` | Add a theme CSS file or directory. Repeatable. |
| `--browser <auto\|chrome\|edge>` | Which installed browser to use. Default: `auto`. Firefox isn't supported. |
| `--browser-path <path>` | Path to a Chrome / Edge / Chromium executable. Default: `$CHROME_PATH`, or found automatically. |
| `-v, --version` | Show this tool's version and the bundled `marp-to-editable-pptx` version. |
| `-h, --help` | Show help. |
| `-- <args…>` | Pass any remaining arguments straight to Marp CLI's Markdown-to-HTML step. |

Added options:

| Option | Description |
| --- | --- |
| `-f, --force` | Overwrite existing output files. |
| `-u, --update` | Only convert when the output is missing or **older** than its input, its Marp config file, or any `--theme` / `--theme-set` file. Stale outputs are overwritten; up-to-date ones are skipped. |
| `-n, --dry-run` | Show what would be converted, overwritten or skipped, without doing it. |
| `-k, --keep-going` | Continue with the remaining files after an error. By default, the tool stops at the first error. |
| `-r, --recursive` | Include subfolders of directories given as plain arguments (`-I` always does). |
| `--output-dir <dir>` | Write the outputs for the given files into `<dir>`, created if needed. Named files go directly into `<dir>`; a directory argument keeps its layout. |
| `-q, --quiet` | Only print errors. |
| `--verbose` | Same as `--debug`. |

### Examples

```sh
# One file -> slide-deck.pptx alongside it (Marp CLI style)
marp-to-editable-pptx-cli --pptx --pptx-editable slide-deck.md

# The same, with the short command
marp2pptx slide-deck.md

# Choose the output name
marp2pptx slide-deck.md -o handout.pptx

# Several files at once, each beside its source
marp2pptx intro.md part1.md part2.md

# Every deck in a folder tree, into build/ (decks/a/b.md -> build/a/b.pptx)
marp2pptx -I decks/ -o build/

# Re-run the same batch, converting only decks that changed since last time
marp2pptx -u -I decks/ -o build/

# Preview what a batch would do
marp2pptx -n -u -I decks/ -o build/

# Named files into one folder
marp2pptx --output-dir build/ intro.md talks/part1.md

# Replace existing outputs unconditionally
marp2pptx -f slide-deck.md

# Pipes: Markdown on stdin, PPTX on stdout
cat slide-deck.md | marp2pptx -o - > slide-deck.pptx

# Custom theme, HTML and local images enabled
marp2pptx --theme-set themes/ --theme corporate --html --allow-local-files slide-deck.md

# Anything else goes straight to Marp CLI
marp2pptx slide-deck.md -- --engine ./engine.js
```

### Overwrite protection and `--update`

The tool never replaces an existing file unless you ask it to:

| Output file… | default | `--update` | `--force` |
| --- | --- | --- | --- |
| doesn't exist | convert | convert | convert |
| exists, older than input | **error**, not overwritten | overwrite | overwrite |
| exists, newer than input | **error**, not overwritten | skip (up to date) | overwrite |

If `--update` and `--force` are combined, `--update` decides. The `--update` check compares file modification times. The output counts as out of date if it is older than the Markdown file, the Marp config file used for that deck (see below), or any `--theme` / `--theme-set` file given on the command line. Images and other files the deck refers to are not checked, so use `--force` after changing only those.

Outputs are written to a temporary file and then renamed into place. A failed or interrupted conversion (Ctrl+C) therefore never leaves a half-written `.pptx`.

### Marp configuration

Marp settings are handled by Marp CLI. Unless `--config-file` or `--no-config-file` is given, each deck gets the project config file (`.marprc.yml`, `marp.config.js`, a `marp` key in `package.json`, …) found in **the deck's own folder**, or else in the nearest parent folder, stopping at the project root (the nearest folder with a `package.json`). As in the VS Code extension, decks in different folders can therefore use different configs, and the result doesn't depend on the directory you run the command from. Relative paths in a config file, such as `themeSet: ./themes`, resolve from the config file's folder.

Command-line options (`--theme`, `--theme-set`, `--html`, and anything after `--`) are applied on top of the config file. Front-matter directives inside each deck (`theme`, `paginate`, `size`, …) work as usual.

Relative image paths in a deck resolve from the deck's own directory. For Markdown read from standard input, they resolve from the current directory.

### Exit codes

| Code | Meaning |
| --- | --- |
| `0` | Every input was converted or was already up to date. |
| `1` | At least one conversion failed, or an existing output was not overwritten. |
| `2` | Invalid usage, for example bad or unsupported options, or a missing input. |
| `130` | Interrupted. |

## API

As with [Marp CLI's API](https://github.com/marp-team/marp-cli#api), the tool can be used from Node.js when it is installed as a dependency of your project:

```sh
npm install github:danielgjackson/marp-to-editable-pptx-cli
```

`marpCli()` takes an array of command-line arguments (any of those in [Options](#options)) and returns a Promise that resolves to the [exit status](#exit-codes) the command would have produced:

```js
import { marpCli } from 'marp-to-editable-pptx-cli'
// or: const { marpCli } = require('marp-to-editable-pptx-cli')

try {
  const exitStatus = await marpCli(['slide-deck.md', '-o', 'handout.pptx', '--force'])

  if (exitStatus > 0) {
    console.error(`Failure (Exit status: ${exitStatus})`)
  } else {
    console.log('Success')
  }
} catch (err) {
  console.error(err)
}
```

The Promise rejects with an `Error` if the run is stopped before any conversion. For invalid usage (bad or unsupported options, a missing input, …) the error is a `CLIError`, whose `errorCode` is a `CLIErrorCode` value. The codes match the exit codes, so invalid usage is `CLIErrorCode.INVALID_OPTIONS` (`2`):

```js
import { marpCli, CLIError } from 'marp-to-editable-pptx-cli'

try {
  await marpCli(['foobar.md'])
} catch (err) {
  if (err instanceof CLIError) {
    console.error(`CLI Error (Error code: ${err.errorCode}):`, err)
  } else {
    console.error('Unexpected error:', err)
  }
}
```

Failed conversions don't reject. They resolve with exit status `1`, as the command does when one file fails, possibly after others have been converted (see `--keep-going`).

Notes:

- Progress and error messages are written to the process's stderr, as on the command line. Pass `--quiet` to show only errors.
- Relative paths resolve from `process.cwd()`.
- Unlike the command, `marpCli()` doesn't read Markdown from stdin when no input is given. It rejects with a usage error instead of waiting for input. Pass `-` as an input to read stdin explicitly, or `marpCli(argv, { stdin: true })` to use the command's behaviour.
- While conversions run, Ctrl+C (`SIGINT`) and `SIGTERM` cancel them, and the Promise resolves with `130`. These signal handlers are removed when the Promise settles.
- TypeScript declarations are included.

## How it works

1. Marp CLI renders the Markdown to HTML in a temporary file next to the source, so that relative image paths resolve.
2. `marp-to-editable-pptx`'s `generateNativePptx()` opens that HTML in a headless browser, reads the laid-out slides from the DOM, and builds an editable presentation with PptxGenJS.

### Relationship to `marp-to-editable-pptx`

The upstream project is used **directly as an npm (dev) dependency**, installed from GitHub and pinned to a release tag. It is only needed to build `dist/cli.js`, which bundles the parts of it this tool uses:

```json
"marp-to-editable-pptx": "github:KatsuYuzu/marp-to-editable-pptx#v1.4.1"
```

No submodule or fork is needed. Upstream's package doesn't ship compiled JavaScript: its `main` points at a VS Code bundle that isn't committed. Its conversion engine (`src/native-pptx/`) has no dependency on VS Code, though, so this tool imports those TypeScript sources and compiles them with esbuild at build time (`scripts/build.js`). The build also regenerates the in-browser DOM-walker script from upstream's `dom-walker.ts`, as upstream's own build does.

Upstream's runtime libraries (`@marp-team/marp-cli`, `puppeteer-core`, `@puppeteer/browsers`, `pptxgenjs`, `cosmiconfig`) are listed as direct dependencies with the same version ranges. They stay external to the bundle and resolve reliably with any package manager.

To move to a newer upstream release, change the tag in `package.json`, copy over any changed version ranges for the libraries above, and run `npm install && npm test`.

## Development

```sh
npm install          # install dependencies (does not build)
npm run bundle       # rebuild dist/cli.js
npm run typecheck    # type-check src/ (and the upstream sources it imports)
npm test             # rebundle, then run the CLI tests (the end-to-end test needs Chrome/Edge)
node dist/cli.js --help
```

`dist/cli.js` is committed, so installing from GitHub needs no build step. It is both the command and the package's `main` module, and the hand-written `types/index.d.ts` describes its API: keep that file in step when the API changes. Rebundle and commit it whenever `src/` or the upstream tag changes. There is deliberately no `build`, `prepare` or `prepack` script: npm would then "prepare" the package when installing it from git, and with `--global` that preparation step installs into the global folder instead of the clone, so the build fails with `Cannot find module 'esbuild'`.

The source is in `src/`: `cli.ts` is the entry point and the `marpCli()` API, `error.ts` defines `CLIError`, `options.ts` handles argument parsing, `plan.ts` expands inputs and applies the overwrite/update rules, and `convert.ts` runs a single conversion.

## Credits

All of the Markdown-to-editable-PowerPoint conversion is done by [marp-to-editable-pptx](https://github.com/KatsuYuzu/marp-to-editable-pptx) by KatsuYuzu (MIT), built on [Marp CLI](https://github.com/marp-team/marp-cli), [Puppeteer](https://pptr.dev/) and [PptxGenJS](https://gitbrent.github.io/PptxGenJS/).
