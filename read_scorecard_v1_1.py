import win32com.client
import os
import re
from pathlib import Path
from openpyxl import Workbook
import pythoncom

def pptx_to_excel_text(val):
    """
    Standardizes PowerPoint text for Excel.
    Converts PPT carriage returns (\r) to Excel newlines (\n).
    """
    if val is None: return ""
    text = str(val)
    # Convert PowerPoint internal breaks to Excel 'Alt+Enter' breaks
    text = text.replace('\r', '\n')
    # Strip illegal ASCII control characters that crash openpyxl
    illegal_chars = re.compile(r'[\000-\010\013\014\016-\037]')
    return illegal_chars.sub("", text)

def get_shape_tag(shape):
    try:
        name = shape.Name.replace(" ", "_")
        if shape.HasTable: kind = "Table"
        elif shape.HasTextFrame:
            text_val = shape.TextFrame.TextRange.Text.strip()
            kind = "Label" if len(text_val) < 25 else "TextBox"
        else:
            s_type = shape.Type
            if s_type == 13: kind = "Image"
            else: kind = "Shape"
        return f"{kind}_{name}"
    except:
        return "Unknown_Shape"


def process_ppt_template(pptx_path):
    """
    Refactored core logic to work with UI wrapper 
    Takes a Path object, returns the path to the generated Excel file.
    """
    pythoncom.CoInitialize()
    pptx_path = Path(pptx_path)
    
    try:
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")

        # Open the presentation
        abs_path = str(pptx_path.absolute())
        prs = ppt_app.Presentations.Open(abs_path, ReadOnly=True, WithWindow=True)
        slide = prs.Slides(1)
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Content"
        ws.append(["Element_Tag", "Original_Content", "Scorecard_1"])

        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            tag = get_shape_tag(shape) # Assumes get_shape_tag is defined above
            content = ""

            if shape.HasTextFrame:
                # Clean and translate text for Excel
                content = pptx_to_excel_text(shape.TextFrame.TextRange.Text)
            
            elif shape.HasTable:
                rows = []
                for r in range(1, shape.Table.Rows.Count + 1):
                    cols = []
                    for c in range(1, shape.Table.Columns.Count + 1):
                        # Clean and translate each cell
                        cell_text = pptx_to_excel_text(shape.Table.Cell(r, c).Shape.TextFrame.TextRange.Text)
                        cols.append(cell_text)
                    rows.append(" | ".join(cols))
                # Separate rows by double pipe
                content = " || ".join(rows)

            ws.append([tag, content, content])
        # print(f"  ✓ {tag}")  --Debugging if certain elements are not captured
        output_excel = pptx_path.parent / (pptx_path.stem + "_xl_template.xlsx")
        wb.save(str(output_excel))
        
        prs.Close()
        return str(output_excel) # Return the path so the UI knows where the file is
    finally:
        #Option but best practice
        pythoncom.CoUninitialize()
        pass

""" def read_template_to_excel():
    raw_path = input("Enter source PPTX path: ").strip().replace('"', '')
    pptx_path = Path(raw_path)
    
    if not pptx_path.exists():
        print(f"File not found: {pptx_path}")
        return

    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
    except:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")

    prs = ppt_app.Presentations.Open(str(pptx_path.absolute()), ReadOnly=True, WithWindow=True)
    slide = prs.Slides(1)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Content"
    ws.append(["Element_Tag", "Original_Content", "Scorecard_1"])

    print("Mapping shapes and standardizing line breaks...")
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        tag = get_shape_tag(shape)
        content = ""

        if shape.HasTextFrame:
            # Clean and translate text for Excel
            content = pptx_to_excel_text(shape.TextFrame.TextRange.Text)
        
        elif shape.HasTable:
            rows = []
            for r in range(1, shape.Table.Rows.Count + 1):
                cols = []
                for c in range(1, shape.Table.Columns.Count + 1):
                    # Clean and translate each cell
                    cell_text = pptx_to_excel_text(shape.Table.Cell(r, c).Shape.TextFrame.TextRange.Text)
                    cols.append(cell_text)
                rows.append(" | ".join(cols))
            # Separate rows by double pipe
            content = " || ".join(rows)

        ws.append([tag, content, content])
        print(f"  ✓ {tag}")

    output_excel = pptx_path.stem + "_input_v1.xlsx"
    wb.save(output_excel)
    prs.Close()
    print(f"\nSuccess. Mapping saved to {output_excel}")
 """


#Allows the script to run standalone if needed
if __name__ == "__main__":
    path = input("Enter PPTX path: ").strip().replace('"', '')
    result = process_template(path)
    print(f"Done: {result}")