# fast-pptx

Quickly make a PowerPoint presentation from a directory of code snippets, CSV files, Graphviz DOT files, images, PDFs, and URLs. **fast-pptx** adds syntax highlighting to the code snippets, converts the CSV files to tables, renders the DOT files, creates high-resolution images from the PDFs, captures screenshots of the websites, and then adds the content to a PowerPoint presentation.

See the [sample output](includes/README_sample_output.md) produced from the included sample input files.

### Author

Paul Stothard

### Quick start

[Install dependencies and download **fast-pptx**](#install), place your source files into a single directory and then run **fast-pptx**:

```bash
./fast-pptx.sh -i input-directory -o output-directory
```

The source files are used to build content, which is added to a PowerPoint presentation created in the output directory. If code snippets are included in the input directory, then a second presentations is created with the code snippets converted to syntax-highlighted code blocks.

The slides can then be edited in PowerPoint to change the order of slides, add or modify text, adjust font sizes, and choose designs for specific slides (using PowerPoint Designer by choosing **Design > Design Ideas** on the ribbon). The file size can then be reduced using **File > Compress Pictures...**.

To combine the content from the two presentations into a single presentation, open both presentations and then copy and paste slides from one presentation to the other. Click on the **Paste Options** button that appears after pasting and choose **Keep Source Formatting**.

### Supported source file types for the input directory

| Type                   | Filename                                                                                                                     | Converted to                      | PowerPoint content generated                                               |
|------------------------|------------------------------------------------------------------------------------------------------------------------------|-----------------------------------|----------------------------------------------------------------------------|
| code snippet           | \*.*language* for example \*.bash \*.perl \*.python \*.r; use `pandoc --list-highlight-languages` to see supported languages | not converted                     | one slide per code snippet file showing syntax-highlighted code            |
| comma-separated values | \*.csv                                                                                                                       | Markdown table                    | one slide per csv file showing the content as a table                      |
| DOT file for Graphviz  | \*.dot                                                                                                                       | png and resized png               | one slide per dot file showing the rendered graph                          |
| gif file               | \*.gif                                                                                                                       | not converted                     | one slide per gif file showing the gif                                     |
| jpg file               | \*.jpg or jpeg                                                                                                               | png and resized png               | one slide per jpg or jpeg file showing the resized png                     |
| pdf file               | \*.pdf                                                                                                                       | png and resized png               | one slide per pdf file showing the resized png                             |
| png file               | \*.png                                                                                                                       | resized png                       | one slide per png file showing the resized png                             |
| svg file               | \*.svg                                                                                                                       | png and resized png               | one slide per svg file showing the resized png                             |
| tiff file              | \*.tiff                                                                                                                      | png and resized png               | one slide per tiff file showing the resized png                            |
| URLs; one per line     | sites.txt                                                                                                                    | png and resized png for each site | one slide per web site URL showing the resized png screenshot for the site |

### Output directory structure

```
outdir
├── slides.pptx
├── slides.md
├── slides_code_blocks.pptx
├── slides_code_blocks.md
└── includes
    ├── resized
```

### Install

**fast-pptx** requires the following:

* [csv2md](https://github.com/pstaender/csv2md)
* [Graphviz](https://graphviz.org)
* [ImageMagick](https://imagemagick.org)
* [pageres-cli](https://github.com/sindresorhus/pageres-cli)
* [pandoc](https://pandoc.org)
* [poppler](https://poppler.freedesktop.org)
* [svgexport](https://github.com/shakiba/svgexport)

On macOS these can be installed as follows:

```bash
brew install graphviz
brew install imagemagick
brew install node
brew install pandoc
brew install poppler
npm install -g csv2md
npm install -g pageres-cli
npm install -g svgexport
```

Clone the repository and test `fast-pptx.sh`:

```bash
git clone git@github.com:paulstothard/fast-pptx.git
cd fast-pptx/scripts
./fast-pptx.sh -i sample_input -o sample_output
```

Or download a [release](https://github.com/paulstothard/fast-pptx/releases/) and test `fast-pptx.sh`, e.g.:

```bash
unzip fast-pptx-1.0.1.zip
cd fast-pptx-1.0.1/scripts
./fast-pptx.sh -i sample_input -o sample_output
```

Add `fast-pptx.sh` to `PATH` or continue to specify the path to `fast-pptx.sh` to run.

### Command-line options

```
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
   -t, --two-column
      For slides containing images generate additional two-column slides.
   -h, --help
      Show this message.

EXAMPLE:
   fast-pptx.sh -i input_dir -o output_dir 
```
