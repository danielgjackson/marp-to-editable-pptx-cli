#!/usr/bin/env node
// Apply the fixes in patches/ to the upstream `marp-to-editable-pptx` sources
// in node_modules, with patch-package. Run on `npm install` (postinstall) and
// before every bundle (scripts/build.js); applying twice is harmless.
//
// Upstream and patch-package are dev dependencies, only needed to build
// dist/cli.js. When the package is installed to be used (e.g. `npm install
// --global github:...`), they are absent and the bundle is already patched,
// so there is nothing to do: skip quietly rather than fail the install.
'use strict'

const { spawnSync } = require('node:child_process')
const path = require('node:path')

const root = path.resolve(__dirname, '..')

function resolves(id) {
  try {
    require.resolve(id, { paths: [root] })
    return true
  } catch {
    return false
  }
}

if (
  !resolves('marp-to-editable-pptx/package.json') ||
  !resolves('patch-package/package.json')
) {
  process.exit(0)
}

const bin = require.resolve('patch-package/index.js', { paths: [root] })
const result = spawnSync(process.execPath, [bin, '--error-on-fail'], {
  cwd: root,
  stdio: 'inherit',
})
process.exit(result.status ?? 1)
