#!/usr/bin/env bash
set -e
mkdir -p ./sample_output/includes
cp ./scripts/includes/theme.pptx ./sample_output/includes/theme.pptx
cp ./scripts/includes/theme_code_blocks.pptx ./sample_output/includes/theme_code_blocks.pptx
./fast-pptx.sh -f -v -i sample_input -o sample_output
