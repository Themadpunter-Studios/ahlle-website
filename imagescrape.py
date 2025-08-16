from openpyxl import load_workbook
from pathlib import Path
from PIL import Image
import io

# Paths
excel_file = "Appel Level Lists (V2).xlsx"
output_folder = Path("thumbnails")
output_folder.mkdir(exist_ok=True)

# Load workbook
wb = load_workbook(excel_file)
ws = wb.active  # or wb['SheetName']

# Map row number to image
image_map = {}
for img in ws._images:  # internal API
    row_num = img.anchor._from.row + 1  # Excel rows are 1-indexed
    image_map[row_num] = img

# Iterate over rows
for row in range(2, 1000):  # D2:D999, A2:A999
    filename = ws[f"A{row}"].value
    if not filename:
        continue

    img = image_map.get(row)
    if not img:
        continue

    # Convert BytesIO to PIL Image
    image_data = img._data()  # returns BytesIO
    pil_img = Image.open(io.BytesIO(image_data))

    # Save image as PNG
    img_path = output_folder / f"{filename}.png"
    pil_img.save(img_path)
    print(f"Saved {img_path}")
