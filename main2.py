

import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
import code128
import io

# Load the Excel file with the data
input_file = 'barcodes.xlsx'  # Replace with the actual file name
output_file = 'output_with_barcodes.xlsx'  # Replace with the actual file name

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(input_file)

# Create a new Excel workbook and select the active sheet
wb = Workbook()
ws = wb.active

# Set the headers for the Excel file
ws.append(['Data', 'Barcode'])

# Define the size for the barcode cells
barcode_column_width = 30
barcode_row_height = 90

# Set the column width for the barcode column
ws.column_dimensions['B'].width = barcode_column_width

# Generate barcodes for each row in the DataFrame and insert them into the Excel file
for index, row in df.iterrows():
    # Assuming the data to generate barcode is in the first column
    data = str(row[0])
    # Generate barcode image using the code128 library
    barcode_image = code128.image(data)
    # Create an in-memory image
    image_stream = io.BytesIO()
    barcode_image.save(image_stream, format='PNG')
    image_stream.seek(0)
    
    # Insert the data and barcode image into the Excel file
    ws.append([data])
    img = OpenpyxlImage(image_stream)
    cell_address = f'B{index + 2}'  # Calculate the cell address where the image will be placed
    img.anchor = cell_address  # Anchor the image using the cell address as a string
    ws.add_image(img)
    
    # Set the row height to accommodate the barcode image
    ws.row_dimensions[index + 2].height = barcode_row_height

# Save the workbook
wb.save(output_file)

print(f"Barcodes have been generated and inserted into {output_file}")
