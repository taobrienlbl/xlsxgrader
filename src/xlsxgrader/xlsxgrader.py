#!/usr/bin/env python3
""" xlsxgrader - a tool for converting assignment exports from a CSV format to an xlsx format for grading purposes. """
from xlsxgrader import parse_canvas_csv
import argparse
from pathlib import Path

def main_cli():
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

def main_gui():
    """ Provides a drag-and-drop interface for converting files."""
    import tkinter as tk
    from tkinterdnd2 import DND_FILES, TkinterDnD

    def open_saveto_dialog(get_directory=True):
        """ Opens a dialog to save the file to a location."""
        from tkinter import filedialog
        if get_directory:
            # ask for a directory to save the files
            return filedialog.askdirectory()
        else:
            # ask for a file to save the file to
            return filedialog.asksaveasfilename()

    # create the main window
    root = TkinterDnD.Tk()
    root.title("CSV to xlsx Converter")
    root.grid_rowconfigure(1, weight=1, minsize = 250)
    root.grid_columnconfigure(0, weight=1, minsize = 300)

    # create the drop target
    drop_target = tk.Listbox(root)
    drop_target.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
    drop_target.insert(1, "Drop files here")
    drop_target.drop_target_register(DND_FILES)

    # make the drop target expand to fill the window
    root.grid_rowconfigure(1, weight=1)

    # create the drop event
    def drop(event):
        # insert the file names into the list
        drop_target.delete(0, tk.END)
        drop_target.insert(tk.END, event.data)

        # check how many files were dropped by inspecting drop_target
        num_files = drop_target.size()

        # if more than one file was dropped, open a dialogue box to determine the save location
        if num_files > 1:
            get_directory = True
        else:
            get_directory = False
        # open a dialogue box to determine the save location
        out_path = open_saveto_dialog()

        # get the files from the drop target
        files = drop_target.get(0, tk.END)
        print(files)

        for file in files:
            # strip the {} from the file name
            file = file[1:-1]
            question_data = parse_canvas_csv.parse_canvas_csv(file)
            if get_directory:
                # strip the .csv extension and add .xlsx and save to the current directory
                # get the file name from file
                output_file = Path(file).with_suffix('.xlsx').name
            else:
                output_file = out_path
            parse_canvas_csv.save_to_xlsx(question_data, output_file)

        quit()


    # bind the drop event to the drop target
    drop_target.dnd_bind("<<Drop>>", drop)

    # pack the drop target
    drop_target.pack()

    # start the main loop
    root.mainloop()
