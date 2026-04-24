# Re-PPT: Automated Slide Generation

RePPT is a Python-based utility designed to automate the creation of repetitive PowerPoint slides. It uses the native PowerPoint COM API to preserve formatting and slide integrity.

---
# Why Re-PPT?

Most slide automation tools are a pain to set up. They force you to use rigid templates, define custom "tags," or write macros before you can even get started.

Re-PPT takes the opposite approach: design your slides first, then automate.

  * Extraction over Setup: Just build your slide. Re-PPT scans it and pulls every shape and text box into an Excel map for you.
  * Excel as the Engine: Once you have the map, you can use formulas and VLOOKUPs to scale one slide to 100 cases instantly.
  * Total Design Freedom: You don't have to worry about "breaking" the automation if you move a box or change a font. The tool adapts to your layout, keeping the final deck pixel-perfect and brand-compliant.

---

## Repository Structure

* `ui_wrapper_v0_2.py`: The Streamlit-based web interface.
* `read_scorecard_v1_1.py`: Module for mapping PPTX shapes to Excel.
* `generate_scorecards_v1_1.py`: Module for duplicating slides and injecting data.
* `requirements.txt`: List of necessary Python libraries.
* `scorecard_generator.bat`: Windows batch file for one-click execution.

---

## Prerequisites

1. **Operating System**: Windows (Required for COM/win32com support).
2. **Software**: Microsoft PowerPoint (Desktop version) must be installed.
3. **Python**: Python 3.9 or higher installed on the system.

---

## Setup Instructions

### 1. Clone the Repository
Download the project folder or clone the repository to your local machine.

### 2. Install Dependencies
Open a terminal in the project directory and run:
```bash
pip install -r requirements.txt
```

### 3. Launch the Application
You can start the tool in two ways:
* **Option A**: Double-click the `run_app.bat` file.
* **Option B**: Run the following command in your terminal:
  ```bash
  streamlit run ui_wrapper_v0_2.py
  ```

---

## Workflow

### Step 1: Create Mapping
1. Navigate to the **Create Mapping Excel** tab.
2. Upload your single-slide PowerPoint template.
3. Download the generated Excel file. This file contains unique tags for every text box and table on your slide.

### Step 2: Edit Data
1. Open the Excel file.
2. Use "Original_Content" as a reference and create new columns for each additional slide you need.
3. For PPTX tables, use `|` to separate columns and `||` to separate rows. Use `Alt+Enter` for new lines within a single cell.

### Step 3: Generate Deck
1. Navigate to the **Generate Scorecards** tab.
2. Upload the original PPTX template and your updated Excel file.
3. Click generate and download the final presentation.

---
## Sample Data
You can find the test input and output files here: https://drive.google.com/drive/folders/1t3scYa00TVdaIPm7Td3Lhfe056qEX9LW?usp=sharing

## Important Notes

* **File Access**: Ensure the PowerPoint files are not open in another program or marked as "Read Only" during the generation process.
* **Tables**: The tool expects the table structure (number of rows/columns) in the template to match the data provided in Excel.
* **Characters**: Standard newlines are automatically converted between Excel and PowerPoint formats to prevent character encoding errors.

## Version Notes

[1.0.0] - 2026-04-24
* Initial internal release
* Supports PPTX textboxes, shapes and tables
