# Excel-Beautify

Modify the appearance of your Excel files by customizing colors, borders, fonts, and more.

![Excel Beautify Example](https://github.com/mTwR0/Excel-beautify/assets/147711036/467b438a-2012-44a9-b630-d8e5056feabc)

## Output Example

![Output Example](https://github.com/mTwR0/Excel-beautify/assets/147711036/425fff51-cf08-4979-9248-7406f94485e8)

## Usage
Every excel sheet that contains data is changed so that it looks the same .

```python
import openpyxl
from excel_functions import set_width, change_headers_border, change_text_border, change_headers_text, change_normal_text, center_text

excel_file = r"your/excel/location.xlsx"
wb = openpyxl.load_workbook(excel_file)

# Column width configuration
column_width = 20

# Text centering
centrare_text = "DA"

# Normal text configuration
schimba_normal_text = "DA"
normal_text_font = "Arial"
normal_text_color = "666666"
normal_text_fill = 'ffffff'
normal_text_size = 10

# Borders for normal text
normal_text_borders = "DA"
normal_text_border = "FULL"  # EXTERIOR sau FULL
normal_text_border_style = "thick"  # dotted, dashed, double, medium, thick, thin
normal_text_border_color = "0080ff"

# Headers text configuration
schimba_headers_text = "DA"
headers_text_font = "Calibri Light"
headers_text_color = "000000"
headers_text_fill = 'c2c2d6'
headers_text_size = 14

# Borders for headers text
headers_text_borders = "DA"
headers_text_border = "FULL"  # EXTERIOR sau FULL
headers_border_style = "thick"  # dotted, dashed, double, medium, thick, thin
headers_border_color = "4d001f"

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]

    max_row = ws.max_row
    max_column = ws.max_column
    min_row = ws.min_row
    min_col = ws.min_column

    set_width(ws, column_width)
    change_headers_border(ws, headers_text_borders, headers_text_border, headers_border_style, headers_border_color, min_row, min_col, max_column)
    change_text_border(normal_text_borders, normal_text_border, ws, min_row, max_row, min_col, max_column, normal_text_border_style, normal_text_border_color)
    change_headers_text(schimba_headers_text, headers_text_font, headers_text_size, headers_text_color, headers_text_fill, ws, min_row, min_col, max_column)
    change_normal_text(schimba_normal_text, normal_text_font, normal_text_size, normal_text_color, normal_text_fill, min_row, ws, max_row, min_col, max_column)
    center_text(centrare_text, ws, min_row, max_row, min_col, max_column)

wb.save(excel_file)
