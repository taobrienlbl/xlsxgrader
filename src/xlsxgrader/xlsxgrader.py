#!/usr/bin/env python3
""" xlsxgrader - a tool for converting assignment exports from a CSV format to an xlsx format for grading purposes. """
import xlsxgrader.parse_canvas_csv
import argparse
from pathlib import Path

def main():
    parser = argparse.ArgumentParser(description='Convert a CSV file to an xlsx file for grading purposes.')
    # allow multiple files to be converted at once
    parser.add_argument('files', metavar='file', type=str, nargs='+', help='the files to convert')
    parser.add_argument('--output', '-o', type=str, help='the output file name (only works if one file is provided)', default="")
    args = parser.parse_args()

    # check if multiple input files were given and if an output file was given
    num_files = len(args.files)
    if num_files > 1 and args.output != "":
        print("Error: cannot specify an output file when converting multiple files")
        return

    # convert each file
    for file in args.files:
        question_data = parse_canvas_csv.parse_canvas_csv(file)

        # if an output file was specified, use that
        if args.output != "":
            parse_canvas_csv.save_to_xlsx(question_data, args.output)
        else:
            # strip the .csv extension and add .xlsx and save to the current directory
            # get the file name from file
            output_file = Path(file).with_suffix('.xlsx').name
            parse_canvas_csv.save_to_xlsx(question_data, output_file)

if __name__ == "__main__":
    main()
