import sys
import argparse

from os import path, makedirs
from pathlib import Path
from textwrap import dedent
from tqdm.auto import tqdm
from tkinter import filedialog, Button, Tk, Label, StringVar
from tkinter.messagebox import showinfo


from win32com import client


def convert(input_folder, output_folder):
    try:
        powerpoint = client.Dispatch("Powerpoint.Application")
        ppSaveAsPDF = 32  # PDF
        files = sorted(Path(input_folder).glob("[!~]*.pptx*"))

        for pptx_file in tqdm(files):
            pdf_file = Path(output_folder) / (str(pptx_file.stem) + ".pdf")
            doc = powerpoint.Presentations.Open(str(pptx_file), WithWindow=False)
            try:
                doc.SaveAs(str(pdf_file), FileFormat=ppSaveAsPDF)
            except:
                raise
            finally:
                doc.Close()

        powerpoint.Quit()
    except:
        raise


def commandline():

    description = dedent(
        """
    Microsoft Powerpoint required.
    Example Usage:

    Convert all files in target directory:
        pptx-print folder/

    Convert all files in target directory, dumping to a different directory:
        pptx-print input_folder/ output_folder/

    """
    )

    formatter_class = lambda prog: argparse.RawDescriptionHelpFormatter(
        prog, max_help_position=32
    )
    parser = argparse.ArgumentParser(
        description=description, formatter_class=formatter_class
    )
    parser.add_argument(
        "input",
        help="input folder",
    )
    parser.add_argument("output", nargs="?", help="output folder")

    args = parser.parse_args()
    input_folder = path.abspath(args.input)
    output_folder = input_folder if not args.output else path.abspath(args.output)

    if not path.exists(input_folder):
        print(f"Input folder '{input_folder}' not found")
        sys.exit(0)

    if not path.exists(output_folder):
        print(f"Creating output folder '{output_folder}'")
        makedirs(output_folder)

    convert(input_folder, output_folder)


def gui():
    root = Tk()
    root.attributes("-topmost", True)
    root.eval('tk::PlaceWindow . center')

    input_folder = StringVar()
    output_folder = StringVar()

    input_label = Label(root,  textvariable=input_folder)
    output_label = Label(root,  textvariable=output_folder)

    def go_convert():
        input = input_folder.get()
        output = output_folder.get() or input

        convert(input, output)
        showinfo(message='Conversion complete.')

    go = Button(root, text="Go", command=go_convert)

    def select_output_folder():
        folder = filedialog.askdirectory()
        if folder:
            output_folder.set(folder)
            output_label.pack()

    select_output = Button(root, text="Set", command=select_output_folder)

    def select_input_folder():
        folder = filedialog.askdirectory()
        if folder:
            input_folder.set(folder)
            input_label.pack()
            select_output.pack()
            go.pack()

    select_input = Button(root, text="Pick", command=select_input_folder)
    select_input.pack()

    root.mainloop()


def run():
    if sys.platform != "win32":
        raise NotImplementedError("This program requires Microsoft Windows.")

    if len(sys.argv) == 1:
        # parser.print_help()
        gui()
    else:
        commandline()


run()
