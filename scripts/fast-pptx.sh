#!/bin/bash -e

force=false
reprocess=false
two_column=true

# store current directory
current_dir=$(pwd)

# Function to change back to the original directory
function return_to_original_dir() {
    cd "$original_dir"
    echo "Returned to the original directory due to an error."
}

# Set a trap to execute the function on errors
trap return_to_original_dir ERR

function error_exit() {
    echo "${PROGNAME}: ${1:-"Unknown Error"}" 1>&2
    exit 1
}

function usage() {
    echo "
USAGE:
   fast-pptx.sh -i DIR -o DIR [Options]

DESCRIPTION:
   Quickly make a PowerPoint presentation from a directory of URLs, images,
   PDFs, CSV files, and code snippets.

REQUIRED ARGUMENTS:
   -i, --input DIR
      Directory of presentation content.
   -o, --output DIR
      Directory for output files.
OPTIONAL ARGUMENTS:
   -f, --force
      Overwrite existing slides.md and pptx files in output directory.
   -r, --reprocess
      Reprocess input files even if conversion files exist in output directory.
   -s, --single-column
      Only generate single-column slides.
   -h, --help
      Show this message.

EXAMPLE:
   fast-pptx.sh -i input_dir -o output_dir  
"
}

while [ "$1" != "" ]; do
    case $1 in
    -i | --input)
        shift
        input=$1
        ;;
    -o | --output)
        shift
        output=$1
        ;;
    -f | --force)
        force=true
        ;;
    -r | --reprocess)
        reprocess=true
        ;;
    -s | --single-column)
        two_column=false
        ;;
    -h | --help)
        usage
        exit
        ;;
    *)
        usage
        exit 1
        ;;
    esac
    shift
done

if [ -z "$input" ]; then
    error_exit "Please use '-i' to specify an input directory. Use '-h' for help."
fi

if [ -z "$output" ]; then
    error_exit "Please use '-o' to specify an output directory. Use '-h' for help."
fi

function end_test() {
  echo "'Check environment' test failed" >&2
  echo "Warning: some conversion steps may fail." >&2
}

for j in pageres dot mmdc csv2md pdftoppm convert svgexport pandoc; do
  if ! command -v $j &>/dev/null; then
    echo "'$j' is required but not installed." >&2
    end_test
  fi
done

if [ ! -d "${output}" ]; then
  mkdir -p "${output}"
fi

if [ ! -d "${output}/includes" ]; then
  mkdir -p "${output}/includes"
fi

#process urls in file input/sites.txt
#save each html file as png using pageres
if [ -f "${input}/sites.txt" ]; then
  while IFS='' read -r url || [ -n "$url" ]; do
    if [ -z "$url" ]; then
      continue
    fi
    case "$url" in \#*) continue ;; esac
    echo "Generating image for URL '$url'."
    output_name=$(echo "$url" | sed -E 's/[^A-Za-z0-9._-]+/_/g')
    if [ -f "${output}/includes/${output_name}.png" ] && ! $reprocess; then
      echo "'$url' has already been processed--skipping."
      continue
    fi
    #these settings give a final image of width 4485 pixels
    #first remove output file if it exists
    rm -f "${output}/includes/${output_name}.png"
    pageres "$url" 897x1090 --crop --scale=5 --filename="${output}/includes/${output_name}"
  done < "${input}/sites.txt"
fi

#convert dot files to graphs using dot
#dot -Tpdf graph.dot -o graph.pdf
find "${input}" -mindepth 1 -maxdepth 1 -iname "*.dot" -type f -exec ls -rt "{}" + | while IFS= read -r dot; do
  file=$(basename -- "$dot")
  echo "Generating pdf for file '$dot'."
  if [ -f "${output}/includes/${file}.pdf" ] && ! $reprocess; then
    echo "'$dot' has already been processed--skipping."
    continue
  fi
  dot -Tpdf "$dot" -o "$output/includes/${file}.pdf"
done

#convert mmd files to graphs using mmdc
#mmdc -i graph.mmd -o graph.pdf --pdfFit
find "${input}" -mindepth 1 -maxdepth 1 -iname "*.mmd" -type f -exec ls -rt "{}" + | while IFS= read -r mmd; do
  file=$(basename -- "$mmd")
  echo "Generating pdf for file '$mmd'."
  if [ -f "${output}/includes/${file}.pdf" ] && ! $reprocess; then
    echo "'$mmd' has already been processed--skipping."
    continue
  fi
  mmdc -i "$mmd" -o "$output/includes/${file}.pdf" --pdfFit
done

#convert csv files to Markdown using csv2md
#csv2md -p data.csv > output.md
find "${input}" -mindepth 1 -maxdepth 1 -iname "*.csv" -type f -exec ls -rt "{}" + | while IFS= read -r csv; do
  file=$(basename -- "$csv")
  echo "Generating Markdown for file '$csv'."
  if [ -f "${output}/includes/${file}.md" ] && ! $reprocess; then
    echo "'$csv' has already been processed--skipping."
    continue
  fi
  #extend short rows to length of first row
  awk -F, -v OFS="," 'NR==1 {cols=NF} {$1=$1; for (i=NF+1; i <= cols; i++) $i = "."} 1' "$csv" > "${output}/includes/${file}.temp"
  csv2md -p < "${output}/includes/${file}.temp" > "${output}/includes/${file}.md"
  rm -f "${output}/includes/${file}.temp"
done

#convert tsv files to Markdown using csv2md
#csv2md -p --csvDelimiter=$'\t' < data.tsv > output.md
find "${input}" -mindepth 1 -maxdepth 1 -iname "*.tsv" -type f -exec ls -rt "{}" + | while IFS= read -r tsv; do
  file=$(basename -- "$tsv")
  echo "Generating Markdown for file '$tsv'."
  if [ -f "${output}/includes/${file}.md" ] && ! $reprocess; then
    echo "'$tsv' has already been processed--skipping."
    continue
  fi
  #extend short rows to length of first row
  awk -F$'\t' -v OFS=$'\t' 'NR==1 {cols=NF} {$1=$1; for (i=NF+1; i <= cols; i++) $i = "."} 1' "$tsv" > "${output}/includes/${file}.temp"
  csv2md -p --csvDelimiter=$'\t' < "${output}/includes/${file}.temp" > "${output}/includes/${file}.md"
  rm -f "${output}/includes/${file}.temp"
done

#cp additional files that are needed in output/includes
#copy any files that are later processed in "${output}/includes/"
find "${input}" -mindepth 1 -maxdepth 1 -not -iname "sites.txt" -not -iname "*.csv" -not -iname "*.dot" -not -iname ".DS_Store" -not -iname "*.jpeg" -not -iname "*.jpg" -not -iname "*.mmd" -not -iname "*.svg" -not -iname "*.tiff" -not -iname "*.tsv" -type f -exec ls -rt "{}" + | while IFS= read -r include_file; do
  file=$(basename -- "$include_file")
  echo "Copying file '$include_file'."
  if [ -f "${output}/includes/${file}" ] && ! $reprocess; then
    echo "'$include_file' has already been copied--skipping."
    continue
  fi
  cp "$include_file" "${output}/includes/${file}"
done

#convert pdf files to png
find "${output}/includes" -mindepth 1 -maxdepth 1 -iname "*.pdf" -type f -exec ls -rt "{}" + | while IFS= read -r pdf; do
  echo "Generating png for '$pdf'."
  if [ -f "${pdf}-1.png" ] || [ -f "${pdf}-01.png" ] || [ -f "${pdf}-001.png" ] && ! $reprocess; then
    echo "'$pdf' has already been processed--skipping."
    continue
  fi
  pdftoppm -f 1 -l 1 -png "$pdf" "${pdf}" -r 600
done

#convert jpg and jpeg images to png
find "${input}" -mindepth 1 -maxdepth 1 \( -iname \*.jpg -o -iname \*.jpeg \) -type f -exec ls -rt "{}" + | while IFS= read -r jpg; do
  file=$(basename -- "$jpg")
  echo "Generating png for '$jpg'."
  if [ -f "${output}/includes/${file}.png" ] && ! $reprocess; then
    echo "'$jpg' has already been processed--skipping."
    continue
  fi
  convert "$jpg" "${output}/includes/${file}.png"
done

#convert tiff images to png
find "${input}" -mindepth 1 -maxdepth 1 \( -iname \*.tiff \) -type f -exec ls -rt "{}" + | while IFS= read -r tiff; do
  file=$(basename -- "$tiff")
  echo "Generating png for '$tiff'."
  if [ -f "${output}/includes/${file}.png" ] && ! $reprocess; then
    echo "'$tiff' has already been processed--skipping."
    continue
  fi
  convert "$tiff" "${output}/includes/${file}.png"
done

#convert svg images to png
find "${input}" -mindepth 1 -maxdepth 1 -iname "*.svg" -type f -exec ls -rt "{}" + | while IFS= read -r svg; do
  file=$(basename -- "$svg")
  echo "Generating png for '$svg'."
  if [ -f "${output}/includes/${file}.png" ] && ! $reprocess; then
    echo "'$svg' has already been processed--skipping."
    continue
  fi
  #convert "$svg" "${output}/includes/${file}.png"
  SVGEXPORT_TIMEOUT=60 svgexport "$svg" "${output}/includes/${file}.png" 4000:
done

#resize images
#PowerPoint slide is 13.33 inches wide at 16:9 setting
#If images are 150 DPI then that is 2000 pixels in width
#If images are 300 DPI then that is 4000 pixels in width
if [ ! -d "${output}/includes/resized" ]; then
  mkdir -p "${output}/includes/resized"
fi

find "${output}/includes" -mindepth 1 -maxdepth 1 -name "*.png" -type f -exec ls -rt "{}" + | while IFS= read -r png; do
  file=$(basename -- "$png")
  echo "Resizing '$png'."
  if [ -f "${output}/includes/resized/${file}" ] && ! $reprocess; then
    echo "'$png' has already been processed--skipping."
    continue
  fi
  convert "$png" -resize 4000 "${output}/includes/resized/${file}"
done

#copy potx and pptx template files
SOURCE="${BASH_SOURCE[0]}"
while [ -h "$SOURCE" ]; do
  DIR="$( cd -P "$( dirname "$SOURCE" )" >/dev/null 2>&1 && pwd )"
  SOURCE="$(readlink "$SOURCE")"
  [[ $SOURCE != /* ]] && SOURCE="$DIR/$SOURCE"
done
DIR="$( cd -P "$( dirname "$SOURCE" )" >/dev/null 2>&1 && pwd )"

find "${DIR}/includes" -mindepth 1 -maxdepth 1 \( -iname \*.potx -o -iname \*.pptx \) -type f -exec ls -rt "{}" + | while IFS= read -r template; do
  file=$(basename -- "$template")
  echo "Copying '$template'."
  if [ -f "${output}/includes/${file}" ]; then
    echo "'$template' has already been copied--skipping."
    continue
  fi
  cp "$template" "${output}/includes/${file}"
done

markdown_file=slides.md
markdown=${output}/${markdown_file}
markdown_code_blocks_file=slides_code_blocks.md
markdown_code_blocks=${output}/${markdown_code_blocks_file}

if [ -f "${markdown}" ] && ! $force; then
  echo "'${markdown}' has already been created."
  echo "Use '--force' to overwrite."
else

echo "Generating Markdown file '$markdown'."

#generate Markdown output
TITLE=$(cat <<-END
% Presentation title
% Name
% Date
END
)

echo "$TITLE" > "$markdown"
echo -e "" >> "$markdown"

echo "$TITLE" > "$markdown_code_blocks"
echo -e "" >> "$markdown_code_blocks"

SECTION=$(cat <<-END
# Section title
END
)

echo "$SECTION" >> "$markdown"
echo -e "" >> "$markdown"

SINGLE_COLUMN_TEXT=$(cat <<-END
## Slide title

Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.

::: notes

Notes

:::
END
)

echo "$SINGLE_COLUMN_TEXT" >> "$markdown"
echo -e "" >> "$markdown"

if $two_column; then 

TWO_COLUMNS_WITH_TEXT=$(cat <<-END
## Slide title

:::::::::::::: {.columns}

::: {.column width="50%"}

Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.

:::

::: {.column width="50%"}

Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.

:::

::::::::::::::

::: notes

Notes

:::
END
)

echo "$TWO_COLUMNS_WITH_TEXT" >> "$markdown"
echo -e "" >> "$markdown"

fi

SINGLE_BULLETED_LIST=$(cat <<-END
## Slide title

- list item
- list item
- list item

::: notes

Notes

:::
END
)

echo "$SINGLE_BULLETED_LIST" >> "$markdown"
echo -e "" >> "$markdown"

SINGLE_BULLETED_LIST_WITH_INDENTING=$(cat <<-END
## Slide title

- list item
  - list item
  - list item
    - list item
- list item

::: notes

Notes

:::
END
)

echo "$SINGLE_BULLETED_LIST_WITH_INDENTING" >> "$markdown"
echo -e "" >> "$markdown"

SINGLE_ORDERED_LIST_WITH_INDENTING=$(cat <<-END
## Slide title

1. list item
   1. list item
   1. list item
      1. list item
   1. list item
1. list item

::: notes

Notes

:::
END
)

echo "$SINGLE_ORDERED_LIST_WITH_INDENTING" >> "$markdown"
echo -e "" >> "$markdown"

if $two_column; then

TWO_COLUMNS_WITH_LISTS=$(cat <<-END
## Slide title

:::::::::::::: {.columns}

::: {.column width="50%"}

- list item
  - list item
  - list item
    - list item
- list item

:::

::: {.column width="50%"}

1. list item
   1. list item
   1. list item
      1. list item
   1. list item
1. list item

:::

::::::::::::::

::: notes

Notes

:::
END
)

echo "$TWO_COLUMNS_WITH_LISTS" >> "$markdown"
echo -e "" >> "$markdown"

fi

#Generate single-column slide for each image and if $two_column generate two-column slide for each image
find "${output}/includes/resized" -mindepth 1 -maxdepth 1 -iname "*.png" -type f -exec ls -rt "{}" + | while IFS= read -r png; do

  # get the filename without the path
  file=$(basename -- "$png")
  # create the path to the file from in the output folder
  png_in_output="./includes/resized/${file}"

  SINGLE_IMAGE=$(cat <<-END
## Slide title

![]($png_in_output)

::: notes

$png_in_output

:::

END
)

  echo "$SINGLE_IMAGE" >> "$markdown"
  echo -e "" >> "$markdown"

if $two_column; then
  SINGLE_IMAGE=$(cat <<-END
## Slide title

:::::::::::::: {.columns}

::: {.column width="50%"}

- Bullet
- Bullet
- Bullet

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

  echo "$SINGLE_IMAGE" >> "$markdown"
  echo -e "" >> "$markdown"
fi

done

#Generate single-column slide for each image and if $two_column generate two-column slide for each image
find "${output}/includes" -mindepth 1 -maxdepth 1 -iname "*.gif" -type f -exec ls -rt "{}" + | while IFS= read -r gif; do

  # get the filename without the path
  file=$(basename -- "$gif")
  # create the path to the file from in the output folder
  gif_in_output="./includes/${file}"

  SINGLE_IMAGE=$(cat <<-END
## Slide title

![]($gif_in_output)

::: notes

$gif_in_output

:::

END
)

  echo "$SINGLE_IMAGE" >> "$markdown"
  echo -e "" >> "$markdown"

if $two_column; then
  SINGLE_IMAGE=$(cat <<-END
## Slide title

:::::::::::::: {.columns}

::: {.column width="50%"}

- Bullet
- Bullet
- Bullet

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

  echo "$SINGLE_IMAGE" >> "$markdown"
  echo -e "" >> "$markdown"
fi

done

#Generate a slide for each Markdown file
find "${output}/includes" -mindepth 1 -maxdepth 1 -iname "*.md" -type f -exec ls -rt "{}" + | while IFS= read -r md; do

  # get the filename without the path
  file=$(basename -- "$md")
  # create the path to the file from in the output folder
  md_in_output="./includes/${file}"

  text=$(<"$md")
  TABLE=$(cat <<-END
## Slide title

$text

::: notes

$md_in_output

:::

END
)

  echo "$TABLE" >> "$markdown"
  echo -e "" >> "$markdown"

done

#Generate a single-column slide for each code file and if $two_column generate two-column slide for each code file
find "${output}/includes" -mindepth 1 -maxdepth 1 -not -iname "sites.txt" -not -iname "*.csv" -not -iname "*.dot" -not -iname ".DS_Store" -not -iname "*.gif" -not -iname "*.jpeg" -not -iname "*.jpg" -not -iname "*.md" -not -iname "*.mmd" -not -iname "*.pdf" -not -iname "*.png" -not -iname "*.pptx" -not -iname "*.potx" -not -iname "*.svg" -not -iname "*.temp" -not -iname "*.tiff" -not -iname "*.tsv" -type f -exec ls -rt "{}" + | while IFS= read -r code; do
  
  # get the filename without the path
  file=$(basename -- "$code")
  # create the path to the file from in the output folder
  code_in_output="./includes/${file}"
  
  #Skip files larger than 1 KB
  maxsize=1000
  filesize=$(du -k "$code" | cut -f1)
  if (( filesize > maxsize )); then
    echo "$code is too large for code block--skipping"
    continue
  fi
  text=$(<"$code")
  extension="${code##*.}"

  CODE=$(cat <<-END
## Slide title

\`\`\`$extension
$text
\`\`\`

::: notes

$code_in_output

:::

END
)

  echo "$CODE" >> "$markdown_code_blocks"
  echo -e "" >> "$markdown_code_blocks"

if $two_column; then
  CODE=$(cat <<-END
## Slide title

:::::::::::::: {.columns}

::: {.column width="50%"}

- Bullet
- Bullet
- Bullet

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

  echo "$CODE" >> "$markdown_code_blocks"
  echo -e "" >> "$markdown_code_blocks"

fi

done

fi

# change to output directory
cd "$output"

# generate pptx file using standard template
pptx=slides.pptx

echo "Generating pptx file '$pptx'."

# generate pptx if "./includes" contains "theme.pptx"
if [ -f "./includes/theme.pptx" ]; then
  # check if the output file exists and if $force is not true
  if [ -f "${pptx}" ] && ! $force; then
    echo "'${pptx}' has already been created."
    echo "Use '--force' to overwrite."
  else
    echo "Generating pptx file '$pptx'."
    
    # create a string containing the pandoc command
    pandoc_command="pandoc $markdown_file -o $pptx --reference-doc ./includes/theme.pptx"

    # run the command
    eval $pandoc_command

    # write the command to the current folder as a script that can be run later
    # include the shebang line and make the file executable
    echo "#!/bin/bash" > pandoc.sh
    echo $pandoc_command >> pandoc.sh
    chmod +x pandoc.sh

  fi
fi

# generate pptx file using code_blocks template
pptx_code_blocks=slides_code_blocks.pptx

# generate pptx if "./includes" contains "theme_code_blocks.pptx"
if [ -f "./includes/theme_code_blocks.pptx" ]; then
  # check if the output file exists and if $force is not true
  if [ -f "${pptx_code_blocks}" ] && ! $force; then
    echo "'${pptx_code_blocks}' has already been created."
    echo "Use '--force' to overwrite."
  else
    echo "Generating pptx file '$pptx_code_blocks'."

    # create a string containing the pandoc command
    pandoc_command="pandoc $markdown_code_blocks_file --highlight-style zenburn -o $pptx_code_blocks --reference-doc ./includes/theme_code_blocks.pptx"

    # run the command
    eval $pandoc_command

    # write the command to the current folder as a script that can be run later
    # include the shebang line and make the file executable
    # check if pandoc.sh exists and if so append to it
    if [ -f "pandoc.sh" ]; then
      echo $pandoc_command >> pandoc.sh
    else
      echo "#!/bin/bash" > pandoc.sh
      echo $pandoc_command >> pandoc.sh
      chmod +x pandoc.sh
    fi
  fi
fi

# change back to original directory
cd "$current_dir"

echo "Done. Check '$output' for slides."
