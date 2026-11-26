#!/usr/bin/env node
// Build a .pptx from a Marp Markdown file with the programmatic API.
// The output is written beside the input, with a .pptx extension.
// Usage: node example/build.js <input.md>  (needs `npm run bundle` first)
'use strict'

const path = require('node:path')
const { marpCli, CLIError } = require('..')

async function main() {
  const input = process.argv[2]
  if (!input || process.argv.length > 3) {
    console.error('Usage: node example/build.js <input.md>')
    process.exitCode = 2
    return
  }
  const parsed = path.parse(path.resolve(input))
  const output = path.join(parsed.dir, `${parsed.name}.pptx`)

  try {
    const exitStatus = await marpCli([input, '-o', output, '--force'])
    if (exitStatus > 0) {
      console.error(`Failure (Exit status: ${exitStatus})`)
    } else {
      console.log(`Success: ${path.relative(process.cwd(), output)}`)
    }
    process.exitCode = exitStatus
  } catch (err) {
    if (err instanceof CLIError) {
      console.error(`CLI Error (Error code: ${err.errorCode}):`, err.message)
      process.exitCode = err.errorCode
    } else {
      console.error('Unexpected error:', err)
      process.exitCode = 1
    }
  }
}

main()
