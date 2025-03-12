import sys
import argparse

from os import path, makedirs
from pathlib import Path
from textwrap import dedent
from tqdm.auto import tqdm

from win32com import client


def convert(input_folder, output_folder):
    try:
        powerpoint = client.Dispatch("Powerpoint.Application")
        ppSaveAsPDF = 32  # PDF
        files = sorted(Path(input_folder).glob("[!~]*.pptx*"))

        for pptx_file in tqdm(files):
            pdf_file = Path(output_folder) / (str(pptx_file.stem) + ".pdf")
            doc = powerpoint.Presentations.Open(str(pptx_file))
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


def run():
    if sys.platform != "win32":
        raise NotImplementedError("This program requires Microsoft Windows.")

    if len(sys.argv) == 1:
        # parser.print_help()
        # gui mode
        sys.exit(0)

    commandline()


run()
