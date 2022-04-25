# pdf-marking-collator

Python script which collates PDF comments and markings.

The `collator.py` allows a collection of PDF files with comments to be combined into a base PDF and can perform tasks such as; mark collation and averaging, marker name obfuscation, generation of marking spreadsheets and overriding of PDF marks via a spreadsheet.

An additional `bulk_collator.py` script is provided which is able to automate the usage of `collator.py` for multiple collections of marked documents. It also allows collecting of the individual spreadsheets from `collator.py`.

## Installation

A `requirements.txt` is included which will allow pip to install the required package dependances with `pip install -r requirements.txt`

## Usage & Examples

Some basic usage is documented here with examples given in the `examples` folder.

### Basic Collation

Consider the case of a submitted PDF called `submission.pdf` which has been marked by two markers who have returned PDFs with comments and marks added.

Basic usage requires providing a directory containing the marked PDFs (`examples\basic`) and the base PDF file (`submission.pdf`) like so:

```
python collator.py examples\basic submission.pdf
```

This produces an `output.pdf` which contains all comments merged into the one document with the averaged marks also written to the document. Marker names have been automatically obfuscated.

### Advanced Collation

Todo...

### Basic Bulk Collation

The `bulk_collator.py` tool is used when you have multiple individual collections which need collated.

Basic usage of the bulk collator requires providing a list of directories containing the individual collections of marked PDFs. In this case the base file is assumed to be the name of the directory.

```
python bulk_collator.py examples\bulk_basic\ada examples\bulk_basic\boltzmann examples\bulk_basic\curie
```

This produces an `output.pdf` in each directory. Essentially, the application of the basic collation to the individual directories.

### Advanced Bulk Collation

Todo...

## Motivation

This tool was originally written to improve the marking process of undergraduate physics reports. It was written to satisfy the needs of the tutors and hopefully be useful to others.
