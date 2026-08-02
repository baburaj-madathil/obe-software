# OBE Software (CO-WK & CO-PO Mapping Engine)

An automated **Outcome-Based Education (OBE)** analysis engine powered by **Gemini AI**.

Built for Higher Education Institutions to automate NBA / Washington Accord GAPC V4.0 alignment. It extracts Course Outcomes (COs) from syllabus PDFs, performs Knowledge Profile (WK1–WK9) mappings, computes GAPC V4.0 Performance Indicator matrix strengths, and serves a multi-user interactive web interface.

---

## 🚀 Key Features

- **AI-Powered Extraction**: Generates 5 precise, measurable Course Outcomes from syllabus PDFs using Gemini AI.
- **Knowledge Profile Mapping**: Maps COs against Washington Accord Knowledge Profiles (WK1–WK9) with detailed keyword justifications.
- **GAPC V4.0 Compliance**: Calculates CO-PO mapping strengths using the official formula:

  $$\text{Mapping Strength } X = \left(\frac{\text{Yes PIs}}{\text{Total PIs}}\right) \times 100$$

  | Range                | Strength | Description  |
  |----------------------|----------|--------------|
  | $X = 68\text{–}100\%$ | **3**    | Substantial  |
  | $X = 34\text{–}67\%$  | **2**    | Moderate     |
  | $X = 10\text{–}33\%$  | **1**    | Slight       |
  | $X < 10\%$           | **0**    | No mapping   |

- **Web UI Server**: Built-in FastAPI application with Server-Sent Events (SSE) streaming real-time terminal progress directly to the browser.
- **Multi-User Isolation**: Handles concurrent user requests by generating uniquely hashed file names in separate `input/` and `output/` folders.

---

## 🛠️ Installation

### Option 1: Install Directly from GitHub

```bash
pip install git+https://github.com/baburaj-madathil/obe-software.git
```

### Option 2: Local Editable Installation (Development)

Clone the repository and install it locally in editable mode:

```bash
git clone https://github.com/baburaj-madathil/obe-software.git
cd obe-software
pip install -e .
```

---

## 🚀 How to Run the Program

You can execute the software using any of the three supported workflows:

### Method 1: Web Application Server (Recommended)

Start the built-in FastAPI web server using the dedicated CLI command:

```bash
obe-server
```

Upon launching, the terminal will display your server access details:

```
=================================================================
           OBE Mapping Server Started Successfully
=================================================================
  Access URL:    http://192.168.1.50:8756
  Input Folder:  /path/to/obe-software/input
  Output Folder: /path/to/obe-software/output
=================================================================
  Press 'q' then ENTER in this window at any time to stop.
=================================================================
```

#### Steps to Use the Web Interface

1. Open the Access URL (e.g., `http://192.168.1.50:8756`) in your web browser.
2. Enter your Gemini API Key.
3. Select and upload your Syllabus PDF file.
4. Click **Process OBE Mapping**.
5. View real-time terminal execution logs and progress updates directly on the page.
6. Once complete, click **Download CO-PO Mapping Excel**.

> 💡 **Graceful Shutdown**: Type `q` and press ENTER in your terminal window at any time to safely shut down the server.

### Method 2: Command Line Interface (CLI)

Run mapping pipelines directly from your terminal.

#### Step 1: Extract CO-WK Mapping from Syllabus PDF

```bash
obe-mapper wk syllabus.pdf --model gemini-2.5-flash --api-key YOUR_GEMINI_API_KEY
```

#### Step 2: Generate GAPC V4.0 CO-PO Mapping from Excel

```bash
obe-mapper po CO_WK_Mapping.xlsx --model gemini-2.5-flash --api-key YOUR_GEMINI_API_KEY
```

### Method 3: Python API

Integrate the mapping engine directly into your custom Python scripts:

```python
from obe import generate_co_wk_excel, generate_co_po_mapping

API_KEY = "YOUR_GEMINI_API_KEY"
MODEL_NAME = "gemini-2.5-flash"

# Step 1: Extract COs and Map Knowledge Profiles (WK1-WK9)
generate_co_wk_excel(
    pdf_path="Syllabus.pdf",
    model_name=MODEL_NAME,
    api_key=API_KEY,
    output_excel_path="CO_WK_Mapping.xlsx"
)

# Step 2: Generate GAPC V4.0 CO-PO Mapping Matrix Report
generate_co_po_mapping(
    model_name=MODEL_NAME,
    api_key=API_KEY,
    input_excel="CO_WK_Mapping.xlsx",
    output_excel="CO_PO_Mapping.xlsx"
)
```
```
