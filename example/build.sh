#!/bin/sh
# Build a .pptx from a Marp Markdown file with the command-line tool.
# The output is written beside the input, with a .pptx extension.
# Usage: ./example/build.sh <input.md> [tool options...]  (needs `npm run bundle` first)
# Extra arguments are passed to the tool, e.g. ./example/build.sh example/example-plain.md --debug
set -e

if [ $# -lt 1 ]; then
  echo "Usage: $0 <input.md> [tool options...]" >&2
  exit 2
fi

INPUT="$1"
shift
case "$(basename "$INPUT")" in
  ?*.*) OUTPUT="${INPUT%.*}.pptx" ;;   # replace the extension
  *)    OUTPUT="$INPUT.pptx" ;;        # no extension: append one
esac

DIR="$(cd "$(dirname "$0")" && pwd)"

node "$DIR/../dist/cli.js" "$INPUT" -o "$OUTPUT" --force "$@"
