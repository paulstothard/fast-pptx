# fast-pptx

Quickly make a PowerPoint presentation from a directory of code snippets, CSV files, TSV files, Graphviz DOT files, Mermaid mmd files, images, PDFs, and URLs. **fast-pptx** adds syntax highlighting to the code snippets, converts the CSV and TSV files to tables, renders the DOT and mmd files, creates high-resolution images from the PDFs, captures screenshots of the websites, and then adds the content to a PowerPoint presentation.

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

| Type                  | Filename        | Converted to                      | PowerPoint Content Generated                                               |
|-----------------------|-----------------|-----------------------------------|----------------------------------------------------------------------------|
| Code Snippet          | *.<language>    | Not Converted                     | One slide per code snippet file showing syntax-highlighted code            |
| CSV File              | *.csv           | Markdown Table                    | One slide per CSV file showing the content as a table                      |
| Dot File for Graphviz | *.dot           | PNG and Resized PNG               | One slide per DOT file showing the rendered graph                          |
| GIF File              | *.gif           | Not Converted                     | One slide per GIF file showing the GIF                                     |
| JPG File              | *.jpg or *.jpeg | PNG and Resized PNG               | One slide per JPG or JPEG file showing the resized PNG                     |
| MMD File for Mermaid  | *.mmd           | PNG and Resized PNG               | One slide per MMD file showing the rendered graph                          |
| PDF File              | *.pdf           | PNG and Resized PNG               | One slide per PDF file showing the resized PNG                             |
| PNG File              | *.png           | Resized PNG                       | One slide per PNG file showing the resized PNG                             |
| SVG File              | *.svg           | PNG and Resized PNG               | One slide per SVG file showing the resized PNG                             |
| TIFF File             | *.tiff          | PNG and Resized PNG               | One slide per TIFF file showing the resized PNG                            |
| TSV File              | *.tsv           | Markdown Table                    | One slide per TSV file showing the content as a table                      |
| URLs (one per line)   | sites.txt       | PNG and Resized PNG for Each Site | One slide per web site URL showing the resized PNG screenshot for the site |

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
* [mermaid-cli](https://github.com/mermaid-js/mermaid-cli)
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
npm install -g mermaid.cli
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
