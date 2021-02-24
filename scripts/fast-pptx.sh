#!/bin/bash -e

#to see supported syntax highlighting
#pandoc --list-highlight-languages

overwrite=false

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
   --overwrite
      Overwrite existing output.
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
    --overwrite)
        overwrite=true
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
  echo "Generating image for URL '$url'"
  output_name=$(echo "$url" | sed -E 's/[^A-Za-z0-9._-]+/_/g')
  if [ -f "${output}/includes/${output_name}.png" ] && ! $overwrite; then
    echo "$url has already been processed--skipping"
    continue
  fi
  #these settings give a final image of width 4485 pixels
  pageres "$url" 897x1090 --crop --scale=5 --filename="${output}/includes/${output_name}"
done < "${input}/sites.txt"

#convert dot files to graphs using graphviz
#dot -Tpdf graph2.dot -o graph2.pd
find "${input}" -mindepth 1 -maxdepth 1 -iname "*.dot" -type f | while IFS= read -r dot; do
  file=$(basename -- "$dot")
  extension="${file##*.}"
  filename="${file%.*}" 
  echo "Generating image for file '$dot'"
  if [ -f "${output}/includes/${file}.pdf" ] && ! $overwrite; then
    echo "'$dot' has already been processed--skipping"
    continue
  fi
  dot -Tpdf "$dot" -o "$output/includes/${file}.pdf"
done

#convert csv files to Markdown using csv2md
#csv2md -p data.csv > output.md
find "${input}" -mindepth 1 -maxdepth 1 -iname "*.csv" -type f | while IFS= read -r csv; do
  file=$(basename -- "$csv")
  extension="${file##*.}"
  filename="${file%.*}" 
  echo "Generating Markdown for file '$csv'"
  if [ -f "${output}/includes/${file}.md" ] && ! $overwrite; then
    echo "'$csv' has already been processed--skipping"
    continue
  fi
  #extend short rows to length of first row
  awk -F, -v OFS="," 'NR==1 {cols=NF} {$1=$1; for (i=NF+1; i <= cols; i++) $i = "."} 1' "$csv" > "${output}/includes/${file}.temp" 
  csv2md -p < "${output}/includes/${file}.temp" > "${output}/includes/${file}.md"
  rm -f "${output}/includes/${file}.temp"
done

#cp pdf files to output/includes
find "${input}" -mindepth 1 -maxdepth 1 -iname "*.pdf" -type f | while IFS= read -r pdf; do
  file=$(basename -- "$pdf")
  extension="${file##*.}"
  filename="${file%.*}"
  echo "Copying file '$pdf'"
  if [ -f "${output}/includes/${file}" ] && ! $overwrite; then
    echo "$'pdf' has already been copied--skipping"
    continue
  fi
  cp "$pdf" "${output}/includes/${file}"
done

#convert pdf files to png
find "${output}/includes" -mindepth 1 -maxdepth 1 -iname "*.pdf" -type f | while IFS= read -r pdf; do
  echo "Generating image for '$pdf'"
  if [ -f "${pdf}-1.png" ] || [ -f "${pdf}-01.png" ] || [ -f "${pdf}-001.png" ] && ! $overwrite; then
    echo "'$pdf' has already been processed--skipping"
    continue
  fi
  pdftoppm -f 1 -l 1 -png "$pdf" "${pdf}" -r 600
done

#convert jpg and jpeg images to png
find "${input}" -mindepth 1 -maxdepth 1 -type f \( -iname \*.jpg -o -iname \*.jpeg \) | while IFS= read -r jpg; do
  file=$(basename -- "$jpg")
  extension="${file##*.}"
  filename="${file%.*}"
  echo "Converting '$jpg'"
  if [ -f "${output}/includes/${file}.png" ] && ! $overwrite; then
    echo "'$jpg' has already been processed--skipping"
    continue
  fi
  convert "$jpg" "${output}/includes/${file}.png"
done

#convert svg images to png
find "${input}" -mindepth 1 -maxdepth 1 -iname "*.svg" -type f | while IFS= read -r svg; do
  file=$(basename -- "$svg")
  extension="${file##*.}"
  filename="${file%.*}"
  echo "Converting '$svg'"
  if [ -f "${output}/includes/${file}.png" ] && ! $overwrite; then
    echo "'$svg' has already been processed--skipping"
    continue
  fi
  convert "$svg" "${output}/includes/${file}.png"
done

#crop images
if [ ! -d "${output}/includes/cropped" ]; then
  mkdir "${output}/includes/cropped"
fi

find "${output}/includes" -mindepth 1 -maxdepth 1 -name "*.png" -type f | while IFS= read -r png; do
  file=$(basename -- "$png")
  extension="${file##*.}"
  filename="${file%.*}"
  echo "Cropping '$png'"
  if [ -f "${output}/includes/cropped/${file}" ] && ! $overwrite; then
    echo "'$png' has already been processed--skipping"
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
  extension="${file##*.}"
  filename="${file%.*}"
  echo "Resizing '$png'"
  if [ -f "${output}/includes/cropped/resized/${file}" ] && ! $overwrite; then
    echo "'$png' has already been processed--skipping"
    continue
  fi
  convert "$png" -resize 4000 "${output}/includes/cropped/resized/${file}"
done

exit

#generate Markdown output
TITLE=$(cat <<-END
% Presentation title
% Name
% Date
END
)

echo "$TITLE" > "$output"
echo -e "" >> "$output"

SECTION=$(cat <<-END
# Section title
END
)

echo "$SECTION" >> "$output"
echo -e "" >> "$output"

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

echo "$SINGLE_BULLETED_LIST" >> "$output"
echo -e "" >> "$output"

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

echo "$SINGLE_BULLETED_LIST_WITH_INDENTING" >> "$output"
echo -e "" >> "$output"

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

echo "$SINGLE_ORDERED_LIST_WITH_INDENTING" >> "$output"
echo -e "" >> "$output"

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

echo "$TWO_COLUMNS_WITH_LISTS" >> "$output"
echo -e "" >> "$output"

#Generate single-column and two-column slide for each image
find "${includes}/cropped/resized" -mindepth 1 -maxdepth 1 -name "*.png" -type f | while IFS= read -r png; do
  SINGLE_IMAGE=$(cat <<-END
## Slide title

![]($png)

::: notes

Speaker notes go here

:::

END
)

  echo "$SINGLE_IMAGE" >> "$output"
  echo -e "" >> "$output"

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

  echo "$SINGLE_IMAGE" >> "$output"
  echo -e "" >> "$output"

done

#Generate a slide for each Markdown file
find "${includes}" -mindepth 1 -maxdepth 1 -name "*.md" -type f | while IFS= read -r md; do
  text=$(<"$md")
  TABLE=$(cat <<-END
## Slide title

$text

::: notes

Speaker notes go here

:::

END
)

  echo "$TABLE" >> "$output"
  echo -e "" >> "$output"

done

#Generate a single-column and two-column slide for each code file
find "${includes}" -mindepth 1 -maxdepth 1 -not -name "*.csv" -not -name "*.dot" -not -name ".DS_Store" -not -name "*.gif" -not -name "*.jpeg" -not -name "*.jpg"  -not -name "*.pdf" -not -name "*.md" -not -name "*.png" -not -name "*.temp" -not -name "urls.txt" -type f | while IFS= read -r code; do
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

  echo "$CODE" >> "$output"
  echo -e "" >> "$output"

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

  echo "$CODE" >> "$output"
  echo -e "" >> "$output"

done

SECTION=$(cat <<-END
# Section Title
END
)

echo "$SECTION" >> "$output"
echo -e "" >> "$output"

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

echo "$SINGLE_BULLETED_LIST" >> "$output"
echo -e "" >> "$output"
