# fast-pptx

Quickly make a PowerPoint presentation from a directory of URLs, images, PDFs, CSV files, and code snippets.


### Author

Paul Stothard


### Quick start

Place your source files into a single directory and then run **fast-pptx**:

```bash
./fast-pptx.sh -i input-directory -o output-directory
```

The source files are used to build content, which is added to a PowerPoint presentation created in the output directory. If PowerPoint templates are included in the input directory, additional presentations are created, one for each tempalte.

The slides can then be edited in PowerPoint to change the order of slides, add or modify text, adjust font sizes, and choose designs for specific slides (using PowerPoint Designer by choosing **Design > Design Ideas** on the ribbon). The file size can then be reduced using **File > Compress Pictures...**.

### Supported source file types for the input directory

| Type                   | Filename                                                 | Converted to                          | PowerPoint content generated                                                                        |
|------------------------|----------------------------------------------------------|---------------------------------------|-----------------------------------------------------------------------------------------------------|
| code snippets          | \*.*language* for example \*.bash \*.perl \*.python \*.r | not converted                         | two slides: one with syntax-highlighted code and one with syntax-highlighted code and a bullet list |
| comma-separated values | \*.csv                                                   | Markdown table                        | one slide per csv file showing the content as a table                                               |
| DOT file for Graphviz  | \*.dot                                                   | pdf and then cropped and resized png  | two slides: one with the png and one with the png and a bullet list                                 |
| gif file               | \*.gif                                                   | not converted                         | two slides: one with the gif and one with the gif and a bullet list                                 |
| jpg file               | \*.jpg or jpeg                                           | cropped and resized png               | two slides: one with the png and one with the png and a bullet list                                 |
| PowerPoint file        | \*.pptx                                                  | not converted                         | used to format the slides                                                                           |
| PowerPoint template    | \*.potx                                                  | not converted                         | used to format the slides                                                                           |
| pdf file               | \*.pdf                                                   | cropped and resized png               | two slides: one with the png and one with the png and a bullet list                                 |
| png file               | \*.png                                                   | cropped and resized png               | two slides: one with the png and one with the png and a bullet list                                 |
| svg file               | \*.svg                                                   | cropped and resized png               | two slides: one with the png and one with the png and a bullet list                                 |
| URLs; one per line     | sites.txt                                                | cropped and resized png for each site | two slides per web site: one with the png and one with the png and a bullet list                    |

### Output directory structure

```
outdir
├── slides.pptx
├── slides.md
└── includes
    ├── cropped
        ├── resized
```

### Docker

The easiest way to run **fast-pptx** is to use the Docker image.

Pull the Docker image:

```bash
docker pull pstothard/fast-pptx
```

Download a sample input directory:

```bash
to-do
```

Run the Docker image and use to create a file:

```bash
docker run --rm -v "$(pwd)":/dir -w /dir pstothard/fast-pptx fast-pptx.sh -i sample_input -o sample_output
```

The `sample_output` directory is written to the current directory on the host system and will contain the resulting PowerPoint presentations.

### Install

**fast-pptx** requires the following:

* [csv2md](https://github.com/pstaender/csv2md)
* [Graphviz](https://graphviz.org)
* [ImageMagick](https://imagemagick.org)
* [pageres-cli](https://github.com/sindresorhus/pageres-cli)
* [pandoc](https://pandoc.org)
* [poppler](https://poppler.freedesktop.org)
* [svgexport](https://github.com/shakiba/svgexport)


Download the script and test data [here](https://github.com/paulstothard/fast-pptx/releases/) or clone the repository:

```bash
git clone git@github.com:paulstothard/fast-pptx.git
cd cgview
```




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
   -h, --help
      Show this message

EXAMPLE:
   fast-pptx.sh -i input_dir -o output_dir  
```

### 


