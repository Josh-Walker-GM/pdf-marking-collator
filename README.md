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

```console
python collator.py examples\basic submission.pdf
```

This produces an `output.pdf` which contains all comments merged into the one document with the averaged marks also written to the document. Marker names have been automatically obfuscated.

### Advanced Collation

A more realisitc approach to the marking process is to have all markers discuss the marks given to a submission and agree upon resolutions to any major discrepancies. This is why the collator tool can generate an `extracted_marks.xlsx` spreadsheet which shows the marks given by each marker to each question. It also includes a plot to allow discrepancies to be seen visually with ease.

Generation of a marking spreadsheet is produced by including the `--generate-spreadsheet` flag like so:

```console
python collator.py examples\advanced submission.pdf --generate-spreadsheet
```

Markers can then make changes to their marks given by simply altering the contents of the spreadsheet and running the collator with the `--use-spreadsheet` flag.

```console
python collator.py examples\advanced submission.pdf --use-spreadsheet
```

This produces the `output.pdf` file with all the collated comments however using the marks as they were provided by the updated spreadsheet. Preventing the need for markers to tinker with their individual PDF marking comments.

### Basic Bulk Collation

The `bulk_collator.py` tool is used when you have multiple individual collections which need collated.

Basic usage of the bulk collator requires providing a list of directories containing the individual collections of marked PDFs. In this case the base file is assumed to be the name of the directory.

```console
python bulk_collator.py examples\bulk_basic\ada examples\bulk_basic\boltzmann examples\bulk_basic\curie
```

This produces an `output.pdf` in each directory. Essentially, the application of the basic collation to the individual directories.

### Advanced Bulk Collation

Again if we wish to follow a more realistic marking approach then we would likely want to have all markers be able to discuss all the marks given across many submissions.

We can have the `bulk_collator.py` tool produce a `combined_extracted_marks.xlsx` spreadsheet gathers the makring spreadsheets for each submission into the one document. This is produced by using the `--generate-combined-spreadsheet` flag like so:

```console
python bulk_collator.py examples\bulk_advanced\ada examples\bulk_advanced\boltzmann examples\bulk_advanced\curie --generate-combined-spreadsheet
```

Once again markers can then edit this combined spreadsheet and then use these updated marks as the ones to be included within the collated PDFs by using a `--use-combined-spreadsheet` flag.

```console
python bulk_collator.py examples\bulk_advanced\ada examples\bulk_advanced\boltzmann examples\bulk_advanced\curie --use-combined-spreadsheet
```

## Motivation

This tool was originally written to improve the marking process of undergraduate physics reports. It was written to satisfy the needs of the tutors at the time and will hopefully prove useful to others in similar situations.
