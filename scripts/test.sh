#!/bin/bash

if [ ! -d "sample_output" ]; then
  mkdir "sample_output"
fi

./fast-pptx.sh -i sample_input -o sample_output
