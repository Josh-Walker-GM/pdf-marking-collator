
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl import load_workbook

import glob
import os
import subprocess
import collections
import statistics
import argparse
import logging

def get_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("directory_list", metavar="directory_list", nargs='+', help="list of directories to collate together")
    parser.add_argument("--generate-individual-spreadsheet", type=bool, default=False, action=argparse.BooleanOptionalAction, help="generate marks spreadsheet for individual collations")
    parser.add_argument("--use-individual-spreadsheet", type=bool, default=False, action=argparse.BooleanOptionalAction, help="use marks spreadsheet to override pdf marks for individual collations")
    parser.add_argument("--generate-combined-spreadsheet", type=bool, default=False, action=argparse.BooleanOptionalAction, help="generate spreadsheet of all collated marks")
    parser.add_argument("--use-combined-spreadsheet", type=bool, default=False, action=argparse.BooleanOptionalAction, help="use combined spreadsheet to override marks from collated pdfs")
    return parser.parse_args()

# def generate_spreadsheet(args, authors: list[str], aliases: list[str], all_marks: list[list[MarkComment]], question_ids: list[str]):
#     wb = Workbook()
#     ws = wb.active
#     wb.save(filename = os.path.join(os.getcwd(), args.input_dir, "extracted_marks.xlsx"))

# def read_spreadsheet(args, authors, question_ids):
#     wb = load_workbook(os.path.join(os.getcwd(), args.input_dir, "extracted_marks.xlsx"))
#     ws = wb.active

#     # read grid of marks
#     overriding_marks: list[list[MarkComment]] = []
#     for col in range(len(authors)):
#         marks: list[MarkComment] = []
#         for row in range(len(question_ids)):
#             mark_comment = MarkComment(None)
#             mark_comment.author = authors[col]
#             mark_comment.question_id = question_ids[row]
#             mark_comment.mark = ws.cell(row+4, col+3).value
#             marks.append(mark_comment)
#         overriding_marks.append(marks)

#     return overriding_marks

def main():

    # set logging format
    logging.basicConfig(format='%(asctime)s: %(message)s', datefmt='%d-%b-%y %H:%M:%S', level=logging.INFO)

    # extract agruments using argparse standard lib
    args = get_arguments()

    # validate against usage of override with and generate spreadsheet features together
    if args.generate_combined_spreadsheet and args.use_combined_spreadsheet:
        logging.error("Cannot use overriding spreadsheet and generate spreadsheet features at the same time!")
        exit(-1)

    # validate against usage of override with and generate individual spreadsheet flags together
    if args.generate_individual_spreadsheet and args.use_individual_spreadsheet:
        logging.error("Cannot use both use and generate individual spreadsheet flags together!")
        exit(-1)

    # validate all collation directories are unique
    if len(args.directory_list) != len(set(args.directory_list)):
        logging.error("Collation directory list cannot contain duplicates!")
        exit(-1)

    # validate all collation directories exist
    for directory in args.directory_list:
        if not os.path.exists(os.path.join(os.getcwd(), directory)):
            logging.error("Collation directory \"{}\" does not exist!".format(os.path.join(os.getcwd(), directory)))
            exit(-1)

    # Dev note: Calling of commands formed from user input is dangerous - should check/sanitise this...
    for directory in args.directory_list:
        logging.info("Collating {}.".format(directory))
        command_string = "python collator.py {} {}.pdf".format(directory, os.path.basename(os.path.normpath(directory)))
        
        if args.generate_individual_spreadsheet:
            command_string = "{} {}".format(command_string, "--generate-spreadsheet")

        if args.use_individual_spreadsheet:
            command_string = "{} {}".format(command_string, "--use-spreadsheet")

        return_code = subprocess.call(command_string, shell=True)
        if return_code != 0:
            logging.error("Collation of \"{}\" failed!".format(directory))
            exit(-1)

if __name__ == "__main__":
    main()
