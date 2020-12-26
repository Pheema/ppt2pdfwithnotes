import comtypes.client
from pathlib import Path

# To get enum values
# Ref: https://stackoverflow.com/questions/52258446/using-file-format-constants-when-saving-powerpoint-presentation-with-comtypes/52315272#52315272
from comtypes.gen import PowerPoint, Office

powerpoint_app = comtypes.client.CreateObject("Powerpoint.Application")
powerpoint_app.Visible = True

input_dir = Path("./ppt")
output_dir = Path("./pdf")
output_dir.mkdir(exist_ok=True)
for file_path in input_dir.glob("*.ppt*"):
    output_file_path = output_dir / f"{file_path.stem}.pdf"

    if output_file_path.exists():
        print(f"{str(output_file_path)} already exists.")
        continue

    presentation = powerpoint_app.Presentations.open(
        str(file_path.absolute()), WithWindow=Office.msoFalse
    )

    print(f"Processing '{str(output_file_path.absolute())}'...")

    try:
        presentation.ExportAsFixedFormat(
            str(output_file_path.absolute()),
            PowerPoint.ppFixedFormatTypePDF,
            Intent=PowerPoint.ppFixedFormatIntentPrint,
            OutputType=PowerPoint.ppPrintOutputNotesPages,  # with notes
        )
    except:
        print(f"Error: exporting failed. '{str(output_file_path.absolute())}'")
    finally:
        presentation.close()

powerpoint_app.quit()
