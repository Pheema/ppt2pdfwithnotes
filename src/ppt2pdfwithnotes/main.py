from pathlib import Path

import win32com.client
from win32com.client import constants as winConstant


def main():
    powerpoint_app = win32com.client.gencache.EnsureDispatch("Powerpoint.Application")
    powerpoint_app.Visible = True
    powerpoint_app.DisplayAlerts = False

    input_dir = Path("./ppt")
    output_dir = Path("./pdf")
    output_dir.mkdir(exist_ok=True)
    for file_path in input_dir.glob("*.ppt*"):
        output_file_path = output_dir / f"{file_path.stem}.pdf"

        if output_file_path.exists():
            print(f"{str(output_file_path)} already exists.")
            continue

        presentation = powerpoint_app.Presentations.Open(str(file_path.absolute()))

        print(f"Processing '{str(output_file_path.absolute())}'...")

        try:
            # PrintRange should be None do to following problem
            # https://stackoverflow.com/questions/17896216/python-win32com-and-powerpoint-exportasfixedformat
            presentation.ExportAsFixedFormat(
                str(output_file_path.absolute()),
                winConstant.ppFixedFormatTypePDF,
                Intent=winConstant.ppFixedFormatIntentPrint,
                OutputType=winConstant.ppPrintOutputNotesPages,  # with notes
                PrintRange=None,
            )
        except:
            print(f"Error: exporting failed. '{str(output_file_path.absolute())}'")
        finally:
            presentation.Close()

    powerpoint_app.Quit()


if __name__ == "__main__":
    main()
