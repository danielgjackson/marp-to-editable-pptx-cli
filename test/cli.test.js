// Black-box tests of argument handling and job planning, via --dry-run (no
// browser needed). Run `npm run bundle` first.
'use strict'

const { test } = require('node:test')
const assert = require('node:assert/strict')
const { spawnSync } = require('node:child_process')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')

const CLI = path.resolve(__dirname, '../dist/cli.js')

function run(args, opts = {}) {
  const r = spawnSync(process.execPath, [CLI, ...args], {
    encoding: 'utf8',
    input: opts.input ?? '',
    cwd: opts.cwd,
  })
  return { code: r.status, out: r.stdout, err: r.stderr }
}

function fixture() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'marp2pptx-test-'))
  fs.mkdirSync(path.join(dir, 'decks/sub/deeper'), { recursive: true })
  fs.mkdirSync(path.join(dir, 'decks/.hidden'))
  fs.writeFileSync(path.join(dir, 'decks/a.md'), '# A')
  fs.writeFileSync(path.join(dir, 'decks/b.markdown'), '# B')
  fs.writeFileSync(path.join(dir, 'decks/notes.txt'), 'not markdown')
  fs.writeFileSync(path.join(dir, 'decks/sub/c.md'), '# C')
  fs.writeFileSync(path.join(dir, 'decks/sub/deeper/d.md'), '# D')
  fs.writeFileSync(path.join(dir, 'decks/.hidden/e.md'), '# E')
  return dir
}

function setMtime(file, secondsAgo) {
  const t = Date.now() / 1000 - secondsAgo
  fs.utimesSync(file, t, t)
}

test('--help and --version', () => {
  assert.match(run(['--help']).out, /^Usage:\n  marp-to-editable-pptx-cli \[options\]/)
  assert.match(
    run(['--version']).out,
    /^marp-to-editable-pptx-cli \S+ \(marp-to-editable-pptx \S+\)/,
  )
})

test('single file defaults to a sibling .pptx', () => {
  const cwd = fixture()
  const r = run(['-n', 'decks/a.md'], { cwd })
  assert.equal(r.code, 0)
  assert.equal(r.err.trim(), `convert: decks/a.md -> ${path.join('decks', 'a.pptx')}`)
})

test('directory input: non-recursive by default, skips dot-dirs', () => {
  const cwd = fixture()
  const r = run(['-n', 'decks'], { cwd })
  assert.equal(r.code, 0)
  assert.match(r.err, /decks\/a\.md -> decks\/a\.pptx/)
  assert.match(r.err, /decks\/b\.markdown -> decks\/b\.pptx/)
  assert.doesNotMatch(r.err, /c\.md|notes|e\.md/)
})

test('recursive directory input mirrors layout into --output-dir', () => {
  const cwd = fixture()
  const r = run(['-n', '-r', '--output-dir', 'out', 'decks'], { cwd })
  assert.equal(r.code, 0)
  const lines = r.err.trim().split('\n')
  assert.deepEqual(lines, [
    'convert: decks/a.md -> out/a.pptx',
    'convert: decks/b.markdown -> out/b.pptx',
    'convert: decks/sub/c.md -> out/sub/c.pptx',
    'convert: decks/sub/deeper/d.md -> out/sub/deeper/d.pptx',
  ])
})

test('explicit files with --output-dir are flattened by basename', () => {
  const cwd = fixture()
  const r = run(['-n', '--output-dir', 'out', 'decks/a.md', 'decks/sub/c.md'], { cwd })
  assert.equal(r.code, 0)
  assert.match(r.err, /decks\/a\.md -> out\/a\.pptx/)
  assert.match(r.err, /decks\/sub\/c\.md -> out\/c\.pptx/)
})

test('existing outputs are protected unless --force or --update', () => {
  const cwd = fixture()
  fs.writeFileSync(path.join(cwd, 'decks/a.pptx'), '')
  assert.match(run(['-n', 'decks/a.md'], { cwd }).err, /skip \(output exists/)
  assert.match(run(['-n', '-f', 'decks/a.md'], { cwd }).err, /convert \(overwrite\)/)

  // Without --dry-run, a refused overwrite is an error.
  const r = run(['decks/a.md'], { cwd })
  assert.equal(r.code, 1)
  assert.match(r.err, /output exists, not overwriting/)
})

test('--update compares modification times', () => {
  const cwd = fixture()
  const md = path.join(cwd, 'decks/a.md')
  const pptx = path.join(cwd, 'decks/a.pptx')
  fs.writeFileSync(pptx, '')

  setMtime(md, 100)
  setMtime(pptx, 10)
  assert.match(run(['-n', '-u', 'decks/a.md'], { cwd }).err, /skip \(up to date\)/)

  setMtime(md, 5)
  assert.match(run(['-n', '-u', 'decks/a.md'], { cwd }).err, /convert \(overwrite\)/)

  // A newer config file also makes the output stale.
  setMtime(md, 100)
  const conf = path.join(cwd, '.marprc.yml')
  fs.writeFileSync(conf, 'html: true\n')
  setMtime(conf, 1)
  assert.match(
    run(['-n', '-u', '-c', '.marprc.yml', 'decks/a.md'], { cwd }).err,
    /convert \(overwrite\)/,
  )
})

test('usage errors exit with code 2', () => {
  const cwd = fixture()
  const cases = [
    [['-o', 'x.pptx', 'decks/a.md', 'decks/sub/c.md'], /--output can only be used with a single input/],
    [['-o', 'x.pptx', '--output-dir', 'out', 'decks/a.md'], /cannot be used together/],
    [['--pdf', 'decks/a.md'], /--pdf is not supported/],
    [['--images', 'png', 'decks/a.md'], /--images is not supported/],
    [['--no-pptx-editable', 'decks/a.md'], /--no-pptx-editable is not supported/],
    [['-w', 'decks/a.md'], /--watch is not supported/],
    [['-I', 'decks', 'decks/a.md'], /either --input-dir or input files/],
    [['-I', 'decks', '-o', '-'], /cannot write to standard output/],
    [['--parallel=x', 'decks/a.md'], /--parallel must be a whole number/],
    [['missing.md'], /input not found/],
    [['-'], /requires --output/],
    [['--browser=firefox', 'decks/a.md'], /--browser must be one of/],
    [['--nope'], /Unknown option/],
  ]
  for (const [args, pattern] of cases) {
    const r = run(args, { cwd })
    assert.equal(r.code, 2, `exit code for ${args.join(' ')}`)
    assert.match(r.err, pattern)
  }
})

test('output colliding with another output or its input is rejected', () => {
  const cwd = fixture()
  fs.writeFileSync(path.join(cwd, 'decks/a.markdown'), '# A2')
  const r = run(['-n', 'decks/a.md', 'decks/a.markdown'], { cwd })
  assert.equal(r.code, 2)
  assert.match(r.err, /would both be written to/)

  const r2 = run(['-n', 'decks/a.md', '-o', 'decks/a.md'], { cwd })
  assert.equal(r2.code, 2)
  assert.match(r2.err, /would overwrite its input/)
})

test('Marp CLI style command lines', () => {
  assert.match(run(['-v']).out, /^marp-to-editable-pptx-cli /)
  const cwd = fixture()

  // --pptx / --pptx-editable are accepted (they are the defaults).
  const r = run(['-n', '--pptx', '--pptx-editable', '--allow-local-files', 'decks/a.md'], { cwd })
  assert.equal(r.code, 0, r.err)
  assert.equal(r.err.trim(), `convert: decks/a.md -> ${path.join('decks', 'a.pptx')}`)

  // -I <dir> -o <dir>: recursive, mirroring the folder structure.
  const batch = run(['-n', '--pptx', '-I', 'decks', '-o', 'out'], { cwd })
  assert.equal(batch.code, 0, batch.err)
  assert.deepEqual(batch.err.trim().split('\n'), [
    'convert: decks/a.md -> out/a.pptx',
    'convert: decks/b.markdown -> out/b.pptx',
    'convert: decks/sub/c.md -> out/sub/c.pptx',
    'convert: decks/sub/deeper/d.md -> out/sub/deeper/d.pptx',
  ])
  // -I without -o writes beside each input.
  assert.match(run(['-n', '-I', 'decks'], { cwd }).err, /decks\/sub\/c\.md -> decks\/sub\/c\.pptx/)

  // -P takes an optional number; a bare -P must not swallow the next file.
  assert.equal(run(['-n', '-P', 'decks/a.md'], { cwd }).code, 0)
  assert.equal(run(['-n', '-P', '2', 'decks/a.md'], { cwd }).code, 0)
  assert.equal(run(['-n', '--no-parallel', 'decks/a.md'], { cwd }).code, 0)

  // --config-file / --no-config-file aliases.
  fs.writeFileSync(path.join(cwd, 'm.yml'), 'html: true\n')
  assert.equal(run(['-n', '--config-file', 'm.yml', 'decks/a.md'], { cwd }).code, 0)
  assert.equal(run(['-n', '--no-config-file', 'decks/a.md'], { cwd }).code, 0)

  // --no-optimize (an extension) is accepted.
  assert.equal(run(['-n', '--no-optimize', 'decks/a.md'], { cwd }).code, 0)
})

test('--update notices a newer Marp config found above the deck', () => {
  const cwd = fixture()
  // The project root (package.json) bounds the upward config search.
  fs.writeFileSync(path.join(cwd, 'package.json'), '{}')
  const md = path.join(cwd, 'decks/sub/c.md')
  const pptx = path.join(cwd, 'decks/sub/c.pptx')
  fs.writeFileSync(pptx, '')
  setMtime(md, 100)
  setMtime(pptx, 10)
  assert.match(run(['-n', '-u', 'decks/sub/c.md'], { cwd }).err, /skip \(up to date\)/)

  const conf = path.join(cwd, 'decks/.marprc.yml')
  fs.writeFileSync(conf, 'html: true\n')
  setMtime(conf, 1)
  assert.match(run(['-n', '-u', 'decks/sub/c.md'], { cwd }).err, /convert \(overwrite\)/)
  // ...unless config lookup is turned off.
  assert.match(
    run(['-n', '-u', '--no-config', 'decks/sub/c.md'], { cwd }).err,
    /skip \(up to date\)/,
  )
})

test('missing --config file is a usage error', () => {
  const cwd = fixture()
  const r = run(['-n', '-c', 'nope.yml', 'decks/a.md'], { cwd })
  assert.equal(r.code, 2)
  assert.match(r.err, /config file not found/)
})

test('programmatic API: marpCli() resolves with the exit status', async () => {
  const { marpCli, CLIError, CLIErrorCode } = require(CLI)
  const cwd = fixture()
  const prevCwd = process.cwd()
  process.chdir(cwd)
  try {
    assert.equal(await marpCli(['-n', '-q', 'decks/a.md']), 0)
    // An existing output is not overwritten: exit status 1, not a rejection.
    fs.writeFileSync('decks/a.pptx', '')
    setMtime('decks/a.md', 10)
    assert.equal(await marpCli(['-q', 'decks/a.md']), 1)
    // Signal handlers are not left installed on the host process.
    assert.equal(process.listenerCount('SIGINT'), 0)
  } finally {
    process.chdir(prevCwd)
  }
  for (const argv of [[], ['missing.md'], ['--pdf', 'x.md']]) {
    await assert.rejects(marpCli(argv), (err) => {
      assert.ok(err instanceof CLIError)
      assert.equal(err.errorCode, CLIErrorCode.INVALID_OPTIONS)
      return true
    })
  }
})

// End-to-end conversion; needs Chrome/Edge, so it is skipped when none is found.
const { detectBrowserPath } = (() => {
  try {
    return require('../dist/cli.js')
  } catch {
    return {}
  }
})()
test('converts decks in parallel into an output directory', { timeout: 180_000 }, (t) => {
  const browser = process.env.CHROME_PATH || detectBrowserPath?.('auto')
  if (!browser) return t.skip('no Chromium-based browser found')
  const cwd = fixture()
  fs.writeFileSync(
    path.join(cwd, 'decks/a.md'),
    '---\nmarp: true\n---\n\n# Hello\n\n- one\n- two\n\n---\n\n## Table\n\n| a | b |\n|---|---|\n| 1 | 2 |\n',
  )
  const r = run(['--pptx', '--pptx-editable', '-P', '2', '-I', 'decks', '-o', 'out'], { cwd })
  assert.equal(r.code, 0, r.err)
  for (const f of ['out/a.pptx', 'out/b.pptx', 'out/sub/c.pptx', 'out/sub/deeper/d.pptx']) {
    const buf = fs.readFileSync(path.join(cwd, f))
    assert.equal(buf.subarray(0, 2).toString(), 'PK', `${f} is a zip/pptx`)
  }
  // No temporary files are left behind.
  const leftovers = fs
    .readdirSync(cwd, { recursive: true })
    .filter((f) => path.basename(String(f)).startsWith('.marp-editable-pptx'))
  assert.deepEqual(leftovers, [])
  // A second --update run converts nothing.
  const again = run(['-u', '-I', 'decks', '-o', 'out'], { cwd })
  assert.equal(again.code, 0, again.err)
  assert.match(again.err, /0 converted, 4 up to date/)
})

test('optimization stores a repeated background image once', { timeout: 180_000 }, async (t) => {
  const browser = process.env.CHROME_PATH || detectBrowserPath?.('auto')
  if (!browser) return t.skip('no Chromium-based browser found')
  const JSZip = require('jszip')
  const cwd = fs.mkdtempSync(path.join(os.tmpdir(), 'marp2pptx-test-'))
  fs.writeFileSync(
    path.join(cwd, 'bg.svg'),
    '<svg xmlns="http://www.w3.org/2000/svg" width="160" height="90">' +
      '<rect width="160" height="90" fill="#cde"/><circle cx="80" cy="45" r="30" fill="#468"/></svg>',
  )
  fs.writeFileSync(
    path.join(cwd, 'deck.md'),
    "---\nmarp: true\nbackgroundImage: url('bg.svg')\n---\n\n# One\n\n---\n\n# Two\n\n---\n\n# Three\n",
  )
  async function inspect(extra) {
    const r = run(['--allow-local-files', '--force', 'deck.md', ...extra], { cwd })
    assert.equal(r.code, 0, r.err)
    const buf = fs.readFileSync(path.join(cwd, 'deck.pptx'))
    const zip = await JSZip.loadAsync(buf)
    const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir)
    const media = names.filter((n) => n.startsWith('ppt/media/'))
    // Every internal relationship target must exist.
    for (const n of names.filter((n) => n.endsWith('.rels'))) {
      const base = path.posix.dirname(path.posix.dirname(n))
      const xml = await zip.file(n).async('string')
      for (const [el, target] of xml.matchAll(/<Relationship\b[^>]*Target="([^"]+)"[^>]*>/g)) {
        if (el.includes('TargetMode="External"')) continue
        const part = path.posix.normalize(path.posix.join(base, target))
        assert.ok(names.includes(part), `${n} points at missing ${target}`)
      }
    }
    return { size: buf.length, media }
  }
  const plain = await inspect(['--no-optimize'])
  const optimized = await inspect([])
  assert.equal(plain.media.length, 3, 'one background copy per slide without optimization')
  assert.equal(optimized.media.length, 1, 'one shared background with optimization')
  assert.match(optimized.media[0], /^ppt\/media\/[0-9a-f]{16}\.png$/)
  assert.ok(optimized.size < plain.size / 2, `${optimized.size} < ${plain.size} / 2`)
})
