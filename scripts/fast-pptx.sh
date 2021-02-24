#!/bin/bash -e

#to see supported syntax highlighting
#pandoc --list-highlight-languages

function error_exit() {
    echo "${PROGNAME}: ${1:-"Unknown Error"}" 1>&2
    exit 1
}

function usage() {
    echo "
USAGE:
   fast-pptx.sh -i DIR

DESCRIPTION:
   Quickly make a PowerPoint presentation from a directory of URLs, images,
   PDFs, CSV files, and code snippets.

REQUIRED ARGUMENTS:
   -i, --includes DIR
      Directory of presentation content.
   -o, --output FILE
      Markdown file to create.
OPTIONAL ARGUMENTS:
   -h, --help
      Show this message

EXAMPLE:
   fast-pptx.sh -i includes  
"
}

while [ "$1" != "" ]; do
    case $1 in
    -i | --includes)
        shift
        includes=$1
        ;;
    -o | --output)
        shift
        output=$1
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

if [ -z "$includes" ]; then
    error_exit "Please use '-i' to specify an includes directory. Use '-h' for help."
fi

if [ -z "$output" ]; then
    error_exit "Please use '-o' to specify an output file. Use '-h' for help."
fi

#process urls in file includes/urls.txt
#save each html file as png using pageres
#sed -E may be OSX specific
while IFS='' read -r url || [ -n "$url" ]; do
  echo "Processing URL '$url'"
  output_name=$(echo "$url" | sed -E 's/[^A-Za-z0-9._-]+/_/g')
  if [ -f "${includes}/${output_name}.pageres.png" ]; then
    echo "$url has already been processed--skipping"
    continue
  fi
  #these settings give a final image of width 4485 pixels
  pageres "$url" 897x1090 --crop --scale=5 --filename="${includes}/${output_name}.pageres"
done < "${includes}/urls.txt"

#convert dot files to graphs using graphviz
#dot -Tpdf graph2.dot -o graph2.pd
find "${includes}" -mindepth 1 -maxdepth 1 -name "*.dot" -type f | while IFS= read -r dot; do
  echo "Processing file '$dot'"
  if [ -f "${dot}.pdf" ]; then
    echo "$dot has already been processed--skipping"
    continue
  fi
  dot -Tpdf "$dot" -o "${dot}.pdf"
done

#convert csv files to Markdown using csv2md
#csv2md -p data.csv > output.md
find "${includes}" -mindepth 1 -maxdepth 1 -name "*.csv" -type f | while IFS= read -r csv; do
  echo "Processing file '$csv'"
  if [ -f "${csv}.md" ]; then
    echo "$csv has already been processed--skipping"
    continue
  fi
  #extend short rows to length of first row
  awk -F, -v OFS="," 'NR==1 {cols=NF} {$1=$1; for (i=NF+1; i <= cols; i++) $i = "."} 1' "$csv" > "${csv}.temp" 
  csv2md -p < "${csv}.temp" > "${csv}.md"
done

#convert pdf files in includes to png
find "${includes}" -mindepth 1 -maxdepth 1 -name "*.pdf" -type f | while IFS= read -r pdf; do
  echo "Processing file '$pdf'"
  if [ -f "${pdf}-1.png" ] || [ -f "${pdf}-01.png" ] || [ -f "${pdf}-001.png" ]; then
    echo "$pdf has already been processed--skipping"
    continue
  fi
  pdftoppm -f 1 -l 1 -png "$pdf" "$pdf" -r 600
done

#convert other jpg and jpeg images to png
find "${includes}" -mindepth 1 -maxdepth 1 -type f \( -iname \*.jpg -o -iname \*.jpeg \) | while IFS= read -r jpg; do
  echo "Processing file '$jpg'"
  if [ -f "${jpg}.png" ]; then
    echo "$jpg has already been processed--skipping"
    continue
  fi
  convert "$jpg" "${jpg}.png"
done

#crop images
if [ ! -d "${includes}/cropped" ]; then
  mkdir "${includes}/cropped"
fi

find "${includes}" -mindepth 1 -maxdepth 1 -name "*.png" -type f | while IFS= read -r png; do
  echo "Processing file '$png'"
  p=$(basename "$png")
  if [ -f "${includes}/cropped/${p}" ]; then
    echo "$png has already been processed--skipping"
    continue
  fi
  convert "$png" -trim -bordercolor White -border 30x30 "${includes}/cropped/${p}"
done

#resize images
#PowerPoint slide is 13.33 inches wide at 16:9 setting
#If images are 150 DPI then that is 2000 pixels in width
#If images are 300 DPI then that is 4000 pixels in width
if [ ! -d "${includes}/cropped/resized" ]; then
  mkdir "${includes}/cropped/resized"
fi

find "${includes}/cropped" -mindepth 1 -maxdepth 1 -name "*.png" -type f | while IFS= read -r png; do
  echo "Processing file '$png'"
  p=$(basename "$png")
  if [ -f "${includes}/cropped/resized/${p}" ]; then
    echo "$png has already been processed--skipping"
    continue
  fi
  convert "$png" -resize 4000 "${includes}/cropped/resized/${p}"
done

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
