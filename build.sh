#!/bin/bash
set -e
cd "$(dirname "$0")"
wat2wasm wat/main.wat -o doc.wasm
echo "Built doc.wasm ($(wc -c < doc.wasm) bytes)"
