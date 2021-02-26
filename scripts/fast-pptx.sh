#!/bin/bash -e

#to see supported syntax highlighting
#pandoc --list-highlight-languages

force=false
reprocess=false

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
   -h, --help
      Show this message

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
  exit 1
}

for j in pageres dot csv2md pdftoppm convert svgexport pandoc; do
  if ! command -v $j &>/dev/null; then
    echo "'$j' is required but not installed." >&2
    end_test
  fi
done

if [ ! -d "${output}" ]; then
  mkdir "${output}"
fi

if [ ! -d "${output}/includes" ]; then
  mkdir "${output}/includes"
fi

#process urls in file input/sites.txt
#save each html file as png using pageres
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
  pageres "$url" 897x1090 --crop --scale=5 --filename="${output}/includes/${output_name}"
done < "${input}/sites.txt"

#convert dot files to graphs using graphviz
#dot -Tpdf graph2.dot -o graph2.pdf
find "${input}" -mindepth 1 -maxdepth 1 -iname "*.dot" -type f | while IFS= read -r dot; do
  file=$(basename -- "$dot")
  echo "Generating pdf for file '$dot'."
  if [ -f "${output}/includes/${file}.pdf" ] && ! $reprocess; then
    echo "'$dot' has already been processed--skipping."
    continue
  fi
  dot -Tpdf "$dot" -o "$output/includes/${file}.pdf"
done

#convert csv files to Markdown using csv2md
#csv2md -p data.csv > output.md
find "${input}" -mindepth 1 -maxdepth 1 -iname "*.csv" -type f | while IFS= read -r csv; do
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

#cp additional files that are needed in output/includes
find "${input}" -mindepth 1 -maxdepth 1 -not -iname "sites.txt" -not -iname "*.csv" -not -iname "*.dot" -not -iname ".DS_Store" -not -iname "*.jpeg" -not -iname "*.jpg" -not -iname "*.svg" -type f | while IFS= read -r include_file; do
  file=$(basename -- "$include_file")
  echo "Copying file '$include_file'."
  if [ -f "${output}/includes/${file}" ] && ! $reprocess; then
    echo "'$include_file' has already been copied--skipping."
    continue
  fi
  cp "$include_file" "${output}/includes/${file}"
done

#convert pdf files to png
find "${output}/includes" -mindepth 1 -maxdepth 1 -iname "*.pdf" -type f | while IFS= read -r pdf; do
  echo "Generating png for '$pdf'."
  if [ -f "${pdf}-1.png" ] || [ -f "${pdf}-01.png" ] || [ -f "${pdf}-001.png" ] && ! $reprocess; then
    echo "'$pdf' has already been processed--skipping."
    continue
  fi
  pdftoppm -f 1 -l 1 -png "$pdf" "${pdf}" -r 600
done

#convert jpg and jpeg images to png
find "${input}" -mindepth 1 -maxdepth 1 -type f \( -iname \*.jpg -o -iname \*.jpeg \) | while IFS= read -r jpg; do
  file=$(basename -- "$jpg")
  echo "Generating png for '$jpg'."
  if [ -f "${output}/includes/${file}.png" ] && ! $reprocess; then
    echo "'$jpg' has already been processed--skipping."
    continue
  fi
  convert "$jpg" "${output}/includes/${file}.png"
done

#convert svg images to png
find "${input}" -mindepth 1 -maxdepth 1 -iname "*.svg" -type f | while IFS= read -r svg; do
  file=$(basename -- "$svg")
  echo "Generating png for '$svg'."
  if [ -f "${output}/includes/${file}.png" ] && ! $reprocess; then
    echo "'$svg' has already been processed--skipping."
    continue
  fi
  #convert "$svg" "${output}/includes/${file}.png"
  SVGEXPORT_TIMEOUT=60 svgexport "$svg" "${output}/includes/${file}.png" 4000:
done

#crop images
if [ ! -d "${output}/includes/cropped" ]; then
  mkdir "${output}/includes/cropped"
fi

find "${output}/includes" -mindepth 1 -maxdepth 1 -name "*.png" -type f | while IFS= read -r png; do
  file=$(basename -- "$png")
  echo "Cropping '$png'."
  if [ -f "${output}/includes/cropped/${file}" ] && ! $reprocess; then
    echo "'$png' has already been processed--skipping."
    continue
  fi
  convert "$png" -trim -bordercolor White -border 30x30 "${output}/includes/cropped/${file}"
done

#resize images
#PowerPoint slide is 13.33 inches wide at 16:9 setting
#If images are 150 DPI then that is 2000 pixels in width
#If images are 300 DPI then that is 4000 pixels in width
if [ ! -d "${output}/includes/cropped/resized" ]; then
  mkdir "${output}/includes/cropped/resized"
fi

find "${output}/includes/cropped" -mindepth 1 -maxdepth 1 -name "*.png" -type f | while IFS= read -r png; do
  file=$(basename -- "$png")
  echo "Resizing '$png'."
  if [ -f "${output}/includes/cropped/resized/${file}" ] && ! $reprocess; then
    echo "'$png' has already been processed--skipping."
    continue
  fi
  convert "$png" -resize 4000 "${output}/includes/cropped/resized/${file}"
done

markdown=${output}/slides.md

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

SECTION=$(cat <<-END
# Section title
END
)

echo "$SECTION" >> "$markdown"
echo -e "" >> "$markdown"

SINGLE_BULLETED_LIST=$(cat <<-END
## Slide title

Single bulleted list:

- list item
- list item
- list item

::: notes

Speaker notes go here

:::
END
)

echo "$SINGLE_BULLETED_LIST" >> "$markdown"
echo -e "" >> "$markdown"

SINGLE_BULLETED_LIST_WITH_INDENTING=$(cat <<-END
## Slide title

Single bulleted list:

- list item
  - list item
  - list item
    - list item
- list item

::: notes

Speaker notes go here

:::
END
)

echo "$SINGLE_BULLETED_LIST_WITH_INDENTING" >> "$markdown"
echo -e "" >> "$markdown"

SINGLE_ORDERED_LIST_WITH_INDENTING=$(cat <<-END
## Slide title

Single ordered list:

1. list item
   1. list item
   1. list item
      1. list item
   1. list item
1. list item

::: notes

Speaker notes go here

:::
END
)

echo "$SINGLE_ORDERED_LIST_WITH_INDENTING" >> "$markdown"
echo -e "" >> "$markdown"

TWO_COLUMNS_WITH_LISTS=$(cat <<-END
## Slide title

:::::::::::::: {.columns}

::: {.column width="50%"}

Left column:

- list item
  - list item
  - list item
    - list item
- list item

:::

::: {.column width="50%"}

Right column:

1. list item
   1. list item
   1. list item
      1. list item
   1. list item
1. list item

:::

::::::::::::::

::: notes

Speaker notes go here

:::
END
)

echo "$TWO_COLUMNS_WITH_LISTS" >> "$markdown"
echo -e "" >> "$markdown"

#Generate single-column and two-column slide for each image
find "${output}/includes/cropped/resized" -mindepth 1 -maxdepth 1 -iname "*.png" -type f | while IFS= read -r png; do
  SINGLE_IMAGE=$(cat <<-END
## Slide title

![]($png)

::: notes

Speaker notes go here

:::

END
)

  echo "$SINGLE_IMAGE" >> "$markdown"
  echo -e "" >> "$markdown"

  SINGLE_IMAGE=$(cat <<-END
## Slide title

:::::::::::::: {.columns}

::: {.column width="50%"}

Left column:

- Bullet
- Bullet
- Bullet

:::

::: {.column width="50%"}

![]($png)

:::

::::::::::::::

::: notes

Speaker notes go here

:::

END
)

  echo "$SINGLE_IMAGE" >> "$markdown"
  echo -e "" >> "$markdown"

done

#Generate single-column and two-column slide for each image
find "${output}/includes" -mindepth 1 -maxdepth 1 -iname "*.gif" -type f | while IFS= read -r gif; do
  SINGLE_IMAGE=$(cat <<-END
## Slide title

![]($gif)

::: notes

Speaker notes go here

:::

END
)

  echo "$SINGLE_IMAGE" >> "$markdown"
  echo -e "" >> "$markdown"

  SINGLE_IMAGE=$(cat <<-END
## Slide title

:::::::::::::: {.columns}

::: {.column width="50%"}

Left column:

- Bullet
- Bullet
- Bullet

:::

::: {.column width="50%"}

![]($gif)

:::

::::::::::::::

::: notes

Speaker notes go here

:::

END
)

  echo "$SINGLE_IMAGE" >> "$markdown"
  echo -e "" >> "$markdown"

done

#Generate a slide for each Markdown file
find "${output}/includes" -mindepth 1 -maxdepth 1 -iname "*.md" -type f | while IFS= read -r md; do
  text=$(<"$md")
  TABLE=$(cat <<-END
## Slide title

$text

::: notes

Speaker notes go here

:::

END
)

  echo "$TABLE" >> "$markdown"
  echo -e "" >> "$markdown"

done

#Generate a single-column and two-column slide for each code file
find "${output}/includes" -mindepth 1 -maxdepth 1 -not -iname "sites.txt" -not -iname "*.csv" -not -iname "*.dot" -not -iname ".DS_Store" -not -iname "*.gif" -not -iname "*.jpeg" -not -iname "*.jpg" -not -iname "*.md" -not -iname "*.pdf" -not -iname "*.png" -not -iname "*.pptx" -not -iname "*.potx" -not -iname "*.svg" -not -iname "*.temp" -type f | while IFS= read -r code; do
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

Speaker notes go here

:::

END
)

  echo "$CODE" >> "$markdown"
  echo -e "" >> "$markdown"

    CODE=$(cat <<-END
## Slide title

:::::::::::::: {.columns}

::: {.column width="50%"}

\`\`\`$extension
$text
\`\`\`

:::

::: {.column width="50%"}

Right column:

- Bullet
- Bullet
- Bullet

:::

::::::::::::::

::: notes

Speaker notes go here

:::

END
)

  echo "$CODE" >> "$markdown"
  echo -e "" >> "$markdown"

done

SECTION=$(cat <<-END
# Section Title
END
)

echo "$SECTION" >> "$markdown"
echo -e "" >> "$markdown"

SINGLE_BULLETED_LIST=$(cat <<-END
## Slide title

Single bulleted list:

- list item
- list item
- list item

::: notes

Speaker notes go here

:::
END
)

echo "$SINGLE_BULLETED_LIST" >> "$markdown"

fi

pptx=${output}/slides.pptx

if [ -f "${pptx}" ] && ! $force; then
  echo "'${pptx}' has already been created."
  echo "Use '--force' to overwrite."
else

echo "Generating pptx file '$pptx'."

pandoc "$markdown" -o "$pptx"

fi

#Generate an additional slide set for each template
find "${output}/includes" -mindepth 1 -maxdepth 1 -type f \( -iname \*.potx -o -iname \*.pptx \) | while IFS= read -r template; do
  file=$(basename -- "$template")
  extension="${file##*.}"
  filename="${file%.*}"

  pptx=${output}/slides_${filename}.pptx
  echo "Generating pptx using template '$template'."
  if [ -f "${pptx}" ] && ! $force; then
    echo "'${pptx}' has already been created."
    echo "Use '--force' to overwrite."
    continue
  fi
  pandoc "$markdown" -o "$pptx" --reference-doc "$template"
done

echo "Done. Check '$output' for slides."
