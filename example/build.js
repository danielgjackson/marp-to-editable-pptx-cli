#!/usr/bin/env node
// Build example.pptx from example.md with the programmatic API.
// Run from anywhere: node example/build.js  (needs `npm run bundle` first)
'use strict'

const path = require('node:path')
const { marpCli, CLIError } = require('..')

const input = path.join(__dirname, 'example.md')
const output = path.join(__dirname, 'example.pptx')

async function main() {
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
