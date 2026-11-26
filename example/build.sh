#!/bin/sh
# Build example.pptx from example.md with the command-line tool.
# Run from anywhere: ./example/build.sh  (needs `npm run bundle` first)
# Extra arguments are passed to the tool, e.g. ./example/build.sh --debug
set -e

DIR="$(cd "$(dirname "$0")" && pwd)"

node "$DIR/../dist/cli.js" "$DIR/example.md" -o "$DIR/example.pptx" --force "$@"
