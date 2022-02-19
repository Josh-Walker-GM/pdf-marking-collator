# pdf-marking-collator
A python script to collate annotations from multiple PDF files into one.

## Usage
Within the directory of the script create a subdirectory which contains the annotated copies of the PDF - eg "sample_dir". Include within that subdirectory an original copy of the unannotated PDF - eg "sample.pdf". 
```
python collator.py sample_dir sample.pdf
```
The following will produce an "output.pdf" which contains a fully annotated copy of "sample.pdf". This example is included within the repo.

## Mark tallying
Annotations which start with "!#" are treated different. They are assumed to correspond to marking annotations which fulfil the following format "!# \[question-id\] \[mark\]" where "question-id" corresponds to some identifier (eg "1", "2a", "3iii") and the mark is assumed to be a numerical value.

All marks with the same question id are averaged together and the result annotated on the output pdf with the total mark also included as an annotation.

--

This was orginally written as a tool for collating marked undergraduate assignments.
