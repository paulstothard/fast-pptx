#!/usr/bin/env bash

# Enable xtrace if the DEBUG environment variable is set
if [[ ${DEBUG-} =~ ^1|yes|true$ ]]; then
  set -o xtrace # Trace the execution of the script (debug)
fi

# Only enable these shell behaviours if we're not being sourced
# Approach via: https://stackoverflow.com/a/28776166/8787985
if ! (return 0 2>/dev/null); then
  # A better class of script...
  set -o errexit  # Exit on most errors (see the manual)
  set -o nounset  # Disallow expansion of unset variables
  set -o pipefail # Use last non-zero exit code in a pipeline
fi

# Enable errtrace or the error trap handler will not work as expected
set -o errtrace # Ensure the error trap handler is inherited

# DESC: Usage help
# ARGS: None
# OUTS: None
function script_usage() {
  cat <<EOF
USAGE:
   fast-pptx.sh -i DIR -o DIR [Options]

DESCRIPTION:
   Quickly make a PowerPoint presentation from a directory of code snippets, 
   CSV files, TSV files, Graphviz DOT files, Mermaid mmd files, images, PDFs, 
   and URLs.

REQUIRED ARGUMENTS:
   -i, --input DIR
      Directory of presentation content.
   -o, --output DIR
      Directory for output files.
OPTIONAL ARGUMENTS:
   -cr, --cron
      Run silently unless error is encountered.
   -f, --force
      Overwrite existing Markdown and PowerPoint files in output directory.
   -h, --help
      Display this message.
   -nc, --no-colour
      Disable colour messages.
   -r, --reprocess
      Reprocess input files even if conversion files exist in output directory.
   -v, --verbose
      Display verbose messages.
EOF
}

# DESC: Parameter parser
# ARGS: $@ (optional): Arguments provided to the script
# OUTS: Variables indicating command-line parameters and options
function parse_params() {
  local param
  while [[ $# -gt 0 ]]; do
    param="$1"
    shift
    case $param in
    -i | --input)
      input=$1
      shift
      ;;
    -o | --output)
      # remove trailing slash
      output="${1%/}"
      shift
      ;;
    -cr | --cron)
      cron=true
      ;;
    -f | --force)
      force=true
      ;;
    -h | --help)
      script_usage
      exit 0
      ;;
    -nc | --no-colour)
      no_colour=true
      ;;
    -r | --reprocess)
      reprocess=true
      ;;
    -v | --verbose)
      verbose=true
      ;;
    *)
      script_exit "Invalid parameter was provided: $param" 1
      ;;
    esac
  done
}

# DESC: Main control flow
# ARGS: $@ (optional): Arguments provided to the script
# OUTS: None
function main() {
  trap script_trap_err ERR
  trap script_trap_exit EXIT

  script_init "$@"
  parse_params "$@"
  cron_init
  colour_init
  #lock_init system

  if [ -z "${input-}" ]; then
    script_exit "Please use '-i' to specify an input directory. Use '-h' for help." 2
  fi

  if [ -z "${output-}" ]; then
    script_exit "Please use '-o' to specify an output directory. Use '-h' for help." 2
  fi

  # check dependencies
  check_binary "pageres"
  check_binary "dot"
  check_binary "mmdc"
  check_binary "awk"
  check_binary "csv2md"
  check_binary "pdftoppm"
  check_binary "convert"
  check_binary "svgexport"
  check_binary "pandoc" 1

  if [ ! -d "${output-}" ]; then
    mkdir -p "${output-}"
  fi

  if [ ! -d "${output-}/includes" ]; then
    mkdir -p "${output-}/includes"
  fi

  if [ ! -d "${output}/includes/resized" ]; then
    mkdir -p "${output}/includes/resized"
  fi

  # process urls in file input/sites.txt
  # save each html file as png using pageres
  if [ -f "${input-}/sites.txt" ]; then
    while IFS='' read -r url || [ -n "$url" ]; do
      if [ -z "$url" ]; then
        continue
      fi
      case "$url" in \#*) continue ;; esac
      verbose_print "Generating image for URL '$url'."
      url_hash=$(echo -n "$url" | md5sum | cut -c1-8)  # Generate a hash of the URL and take the first 8 characters
      url_substring=$(echo "$url" | cut -c1-32)  # Take the first 32 characters of the URL
      output_name="${url_hash}_${url_substring}"  # Concatenate the hash and the substring
      output_name=$(echo "$output_name" | sed -E 's/[^A-Za-z0-9._-]+/_/g')  # Replace any non-alphanumeric characters
      output_name=$(echo "$output_name" | cut -c1-50)  # Limit the output name to 50 characters
      if [ -f "${output-}/includes/${output_name}.png" ] && [ -z "${reprocess-}" ]; then
        verbose_print "'$url' has already been processed--skipping."
        continue
      fi
      # these settings give a final image of width 4485 pixels
      # first remove output file if it exists
      rm -f "${output-}/includes/${output_name}.png"
      pageres "$url" 897x1090 --crop --scale=5 --filename="${output-}/includes/${output_name}"
    done <"${input-}/sites.txt"
  fi

  # convert dot files to graphs using dot
  # dot -Tpdf graph.dot -o graph.pdf
  find "${input-}" -mindepth 1 -maxdepth 1 -iname "*.dot" -type f -exec ls -rt "{}" + | while IFS= read -r dot; do
    file=$(basename -- "$dot")
    verbose_print "Generating pdf for file '$dot'."
    if [ -f "${output-}/includes/${file}.pdf" ] && [ -z "${reprocess-}" ]; then
      verbose_print "'$dot' has already been processed--skipping."
      continue
    fi
    dot -Tpdf "$dot" -o "${output-}/includes/${file}.pdf"
  done

  # convert mmd files to graphs using mmdc
  # mmdc -i graph.mmd -o graph.pdf --pdfFit
  find "${input-}" -mindepth 1 -maxdepth 1 -iname "*.mmd" -type f -exec ls -rt "{}" + | while IFS= read -r mmd; do
    file=$(basename -- "$mmd")
    verbose_print "Generating pdf for file '$mmd'."
    if [ -f "${output-}/includes/${file}.pdf" ] && [ -z "${reprocess-}" ]; then
      verbose_print "'$mmd' has already been processed--skipping."
      continue
    fi
    mmdc -i "$mmd" -o "${output-}/includes/${file}.pdf" --pdfFit
  done

  # convert csv files to Markdown using csv2md
  # csv2md -p data.csv > output.md
  find "${input-}" -mindepth 1 -maxdepth 1 -iname "*.csv" -type f -exec ls -rt "{}" + | while IFS= read -r csv; do
    file=$(basename -- "$csv")
    verbose_print "Generating Markdown for file '$csv'."
    if [ -f "${output-}/includes/${file}.md" ] && [ -z "${reprocess-}" ]; then
      verbose_print "'$csv' has already been processed--skipping."
      continue
    fi
    # extend short rows to length of first row
    awk -F, -v OFS="," 'NR==1 {cols=NF} {$1=$1; for (i=NF+1; i <= cols; i++) $i = "."} 1' "$csv" >"${output-}/includes/${file}.temp"
    csv2md -p <"${output-}/includes/${file}.temp" >"${output-}/includes/${file}.md"
    rm -f "${output-}/includes/${file}.temp"
  done

  # convert tsv files to Markdown using csv2md
  # csv2md -p --csvDelimiter=$'\t' < data.tsv > output.md
  find "${input-}" -mindepth 1 -maxdepth 1 -iname "*.tsv" -type f -exec ls -rt "{}" + | while IFS= read -r tsv; do
    file=$(basename -- "$tsv")
    verbose_print "Generating Markdown for file '$tsv'."
    if [ -f "${output-}/includes/${file}.md" ] && [ -z "${reprocess-}" ]; then
      verbose_print "'$tsv' has already been processed--skipping."
      continue
    fi
    # extend short rows to length of first row
    awk -F$'\t' -v OFS=$'\t' 'NR==1 {cols=NF} {$1=$1; for (i=NF+1; i <= cols; i++) $i = "."} 1' "$tsv" >"${output-}/includes/${file}.temp"
    csv2md -p --csvDelimiter=$'\t' <"${output-}/includes/${file}.temp" >"${output-}/includes/${file}.md"
    rm -f "${output-}/includes/${file}.temp"
  done

  # cp additional files that are needed in output/includes
  # copy any files that are later processed in "${output}/includes/"
  find "${input-}" -mindepth 1 -maxdepth 1 -not -iname "sites.txt" -not -iname "*.csv" -not -iname "*.dot" -not -iname ".DS_Store" -not -iname "*.jpeg" -not -iname "*.jpg" -not -iname "*.mmd" -not -iname "*.svg" -not -iname "*.tiff" -not -iname "*.tsv" -type f -exec ls -rt "{}" + | while IFS= read -r include_file; do
    file=$(basename -- "$include_file")
    verbose_print "Copying file '$include_file'."
    if [ -f "${output-}/includes/${file}" ] && [ -z "${reprocess-}" ]; then
      verbose_print "'$include_file' has already been copied--skipping."
      continue
    fi
    cp "$include_file" "${output-}/includes/${file}"
  done

  # convert pdf files to png
  find "${output-}/includes" -mindepth 1 -maxdepth 1 -iname "*.pdf" -type f -exec ls -rt "{}" + | while IFS= read -r pdf; do
    verbose_print "Generating png for '$pdf'."
    if [ -f "${pdf}-1.png" ] || [ -f "${pdf}-01.png" ] || [ -f "${pdf}-001.png" ] && [ -z "${reprocess-}" ]; then
      verbose_print "'$pdf' has already been processed--skipping."
      continue
    fi
    pdftoppm -png "$pdf" "${pdf}" -r 600
  done

  # convert jpg and jpeg images to png
  find "${input-}" -mindepth 1 -maxdepth 1 \( -iname \*.jpg -o -iname \*.jpeg \) -type f -exec ls -rt "{}" + | while IFS= read -r jpg; do
    file=$(basename -- "$jpg")
    verbose_print "Generating png for '$jpg'."
    if [ -f "${output-}/includes/${file}.png" ] && [ -z "${reprocess-}" ]; then
      verbose_print "'$jpg' has already been processed--skipping."
      continue
    fi
    convert "$jpg" "${output-}/includes/${file}.png"
  done

  # convert tiff images to png
  find "${input-}" -mindepth 1 -maxdepth 1 \( -iname \*.tiff \) -type f -exec ls -rt "{}" + | while IFS= read -r tiff; do
    file=$(basename -- "$tiff")
    verbose_print "Generating png for '$tiff'."
    if [ -f "${output-}/includes/${file}.png" ] && [ -z "${reprocess-}" ]; then
      verbose_print "'$tiff' has already been processed--skipping."
      continue
    fi
    convert "$tiff" "${output-}/includes/${file}.png"
  done

  # convert svg images to png
  find "${input-}" -mindepth 1 -maxdepth 1 -iname "*.svg" -type f -exec ls -rt "{}" + | while IFS= read -r svg; do
    file=$(basename -- "$svg")
    verbose_print "Generating png for '$svg'."
    if [ -f "${output-}/includes/${file}.png" ] && [ -z "${reprocess-}" ]; then
      verbose_print "'$svg' has already been processed--skipping."
      continue
    fi
    SVGEXPORT_TIMEOUT=60 svgexport "$svg" "${output-}/includes/${file}.png" 4000:
  done

  # resize images
  # PowerPoint slide is 13.33 inches wide at 16:9 setting
  # If images are 150 DPI then that is 2000 pixels in width
  # If images are 300 DPI then that is 4000 pixels in width
  find "${output-}/includes" -mindepth 1 -maxdepth 1 -name "*.png" -type f -exec ls -rt "{}" + | while IFS= read -r png; do
    file=$(basename -- "$png")
    verbose_print "Resizing '$png'."
    if [ -f "${output-}/includes/resized/${file}" ] && [ -z "${reprocess-}" ]; then
      verbose_print "'$png' has already been processed--skipping."
      continue
    fi
    convert "$png" -resize 4000 "${output}/includes/resized/${file}"
  done

  # copy potx and pptx template files
  find "${script_dir}/includes" -mindepth 1 -maxdepth 1 \( -iname \*.potx -o -iname \*.pptx \) -type f -exec ls -rt "{}" + | while IFS= read -r template; do
    file=$(basename -- "$template")
    verbose_print "Copying '$template'."
    if [ -f "${output-}/includes/${file}" ] && [ -z "${reprocess-}" ]; then
      verbose_print "'$template' has already been copied--skipping."
      continue
    fi
    cp "$template" "${output-}/includes/${file}"
  done

  markdown_file=slides.md
  markdown_code_blocks_file=slides_code_blocks.md

  markdown=${output-}/${markdown_file}
  markdown_code_blocks=${output-}/${markdown_code_blocks_file}

  if [ -f "${markdown}" ] && [ -z "${force-}" ]; then
    script_exit "'${markdown}' has already been created. Use '--force' to overwrite." 2
  fi

  if [ -f "${markdown_code_blocks}" ] && [ -z "${force-}" ]; then
    script_exit "'${markdown_code_blocks}' has already been created. Use '--force' to overwrite." 2
  fi

  verbose_print "Generating Markdown files."

  TITLE=$(
    cat <<-END
% Presentation title
% Name
% Date
END
  )

  echo "$TITLE" >"$markdown"
  echo -e "" >>"$markdown"

  echo "$TITLE" >"$markdown_code_blocks"
  echo -e "" >>"$markdown_code_blocks"

  SECTION=$(
    cat <<-END
# Section title
END
  )

  echo "$SECTION" >>"$markdown"
  echo -e "" >>"$markdown"

  SINGLE_BULLETED_LIST=$(
    cat <<-END
## Slide title

- Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.
- Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.
- Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.

::: notes

Notes

:::
END
  )

  echo "$SINGLE_BULLETED_LIST" >>"$markdown"
  echo -e "" >>"$markdown"

  # generate slides for each image
  find "${output-}/includes/resized" -mindepth 1 -maxdepth 1 -iname "*.png" -type f -exec ls -rt "{}" + | while IFS= read -r png; do

    # get the filename without the path
    file=$(basename -- "$png")
    # create the path to the file from in the output folder
    png_in_output="./includes/resized/${file}"

    SINGLE_IMAGE=$(
      cat <<-END
## Slide title

![]($png_in_output)

::: notes

$png_in_output

:::

END
    )

    echo "$SINGLE_IMAGE" >>"$markdown"
    echo -e "" >>"$markdown"

    SINGLE_IMAGE=$(
      cat <<-END
## Slide title

:::::::::::::: {.columns}

::: {.column width="50%"}

- Lorem ipsum dolor sit.
- Eiusmod tempor incididunt.
- Enim ad minim veniam.

:::

::: {.column width="50%"}

![]($png_in_output)

:::

::::::::::::::

::: notes

$png_in_output

:::

END
    )

    echo "$SINGLE_IMAGE" >>"$markdown"
    echo -e "" >>"$markdown"

  done

  # generate slides for each image
  find "${output-}/includes" -mindepth 1 -maxdepth 1 -iname "*.gif" -type f -exec ls -rt "{}" + | while IFS= read -r gif; do

    # get the filename without the path
    file=$(basename -- "$gif")
    # create the path to the file from in the output folder
    gif_in_output="./includes/${file}"

    SINGLE_IMAGE=$(
      cat <<-END
## Slide title

![]($gif_in_output)

::: notes

$gif_in_output

:::

END
    )

    echo "$SINGLE_IMAGE" >>"$markdown"
    echo -e "" >>"$markdown"

    SINGLE_IMAGE=$(
      cat <<-END
## Slide title

:::::::::::::: {.columns}

::: {.column width="50%"}

- Lorem ipsum dolor sit.
- Eiusmod tempor incididunt.
- Enim ad minim veniam.

:::

::: {.column width="50%"}

![]($gif_in_output)

:::

::::::::::::::

::: notes

$gif_in_output

:::

END
    )

    echo "$SINGLE_IMAGE" >>"$markdown"
    echo -e "" >>"$markdown"

  done

  # generate slides for each Markdown file
  find "${output-}/includes" -mindepth 1 -maxdepth 1 -iname "*.md" -type f -exec ls -rt "{}" + | while IFS= read -r md; do

    # get the filename without the path
    file=$(basename -- "$md")
    # create the path to the file from in the output folder
    md_in_output="./includes/${file}"

    text=$(<"$md")

    TABLE=$(
      cat <<-END
## Slide title

$text

::: notes

$md_in_output

:::

END
    )

    echo "$TABLE" >>"$markdown"
    echo -e "" >>"$markdown"

    TABLE=$(
      cat <<-END
## Slide title

:::::::::::::: {.columns}

::: {.column width="50%"}

- Lorem ipsum dolor sit.
- Eiusmod tempor incididunt.
- Enim ad minim veniam.

:::

::: {.column width="50%"}

$text

:::

::::::::::::::

::: notes

$md_in_output

:::

END
    )

    echo "$TABLE" >>"$markdown"
    echo -e "" >>"$markdown"

  done

  # generate slides for each code file
  find "${output-}/includes" -mindepth 1 -maxdepth 1 -not -iname "sites.txt" -not -iname "*.csv" -not -iname "*.dot" -not -iname ".DS_Store" -not -iname "*.gif" -not -iname "*.jpeg" -not -iname "*.jpg" -not -iname "*.md" -not -iname "*.mmd" -not -iname "*.pdf" -not -iname "*.png" -not -iname "*.pptx" -not -iname "*.potx" -not -iname "*.svg" -not -iname "*.temp" -not -iname "*.tiff" -not -iname "*.tsv" -type f -exec ls -rt "{}" + | while IFS= read -r code; do

    # get the filename without the path
    file=$(basename -- "$code")
    # create the path to the file from in the output folder
    code_in_output="./includes/${file}"

    # skip files larger than 1 KB
    maxsize=1000
    filesize=$(du -k "$code" | cut -f1)
    if ((filesize > maxsize)); then
      verbose_print "$code is too large for code block--skipping"
      continue
    fi
    text=$(<"$code")
    extension="${code##*.}"

    CODE=$(
      cat <<-END
## Slide title

\`\`\`$extension
$text
\`\`\`

::: notes

$code_in_output

:::

END
    )

    echo "$CODE" >>"$markdown_code_blocks"
    echo -e "" >>"$markdown_code_blocks"

    CODE=$(
      cat <<-END
## Slide title

:::::::::::::: {.columns}

::: {.column width="50%"}

- Lorem ipsum dolor sit.
- Eiusmod tempor incididunt.
- Enim ad minim veniam.

:::

::: {.column width="50%"}

\`\`\`$extension
$text
\`\`\`

:::

::::::::::::::

::: notes

$code_in_output

:::

END
    )

    echo "$CODE" >>"$markdown_code_blocks"
    echo -e "" >>"$markdown_code_blocks"

  done

  # generate pptx file using standard template
  pptx=${output-}/slides.pptx

  verbose_print "Generating pptx file '$pptx'."

  # check if the output file exists and if $force is not true
  if [ -f "${pptx}" ] && [ -z "${force-}" ]; then
    script_exit "'${pptx}' has already been created. Use '--force' to overwrite." 2
  fi

  # convert the Markdown file to pptx
  pandoc "$markdown" --resource-path="${output-}" -o "$pptx" --reference-doc "./includes/theme.pptx"

  # create a script that can be used to regenerate the pptx file
  # include the shebang line and make the file executable
  pandoc_command="pandoc $markdown_file -o slides.pptx --reference-doc ./includes/theme.pptx"
  echo "#!/bin/bash" >"${output-}/pandoc.sh"
  echo "$pandoc_command" >>"${output-}/pandoc.sh"
  chmod +x "${output-}/pandoc.sh"

  # generate pptx file using code_blocks template
  pptx_code_blocks=${output-}/slides_code_blocks.pptx

  verbose_print "Generating pptx file '$pptx'."

  # check if the output file exists and if $force is not true
  if [ -f "${pptx_code_blocks}" ] && [ -z "${force-}" ]; then
    script_exit "'${pptx_code_blocks}' has already been created. Use '--force' to overwrite." 2
  fi

  # convert the Markdown file to pptx
  pandoc "$markdown_code_blocks" --resource-path="${output-}" --highlight-style zenburn -o "$pptx_code_blocks" --reference-doc "./includes/theme_code_blocks.pptx"

  # create a script that can be used to regenerate the pptx file
  # include the shebang line and make the file executable
  # check if pandoc.sh exists and if so append to it
  pandoc_command="pandoc $markdown_code_blocks_file --highlight-style zenburn -o slides_code_blocks.pptx --reference-doc ./includes/theme_code_blocks.pptx"

  if [ -f "${output-}/pandoc.sh" ]; then
    echo "$pandoc_command" >>"${output-}/pandoc.sh"
  else
    echo "#!/bin/bash" >"${output-}/pandoc.sh"
    echo "$pandoc_command" >>"${output-}/pandoc.sh"
    chmod +x "${output-}/pandoc.sh"
  fi

  verbose_print "Done. Check '$output' for slides."

}

# shellcheck source=source.sh
source "$(dirname "${BASH_SOURCE[0]}")/source.sh"

# Invoke main with args if not sourced
# Approach via: https://stackoverflow.com/a/28776166/8787985
if ! (return 0 2>/dev/null); then
  main "$@"
fi
