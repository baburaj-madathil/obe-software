# OBE Software (CO-WK & CO-PO Mapping Engine)

An automated **Outcome-Based Education (OBE)** analysis engine powered by **Gemini AI**. Built for Higher Education Institutions to automate NBA / Washington Accord GAPC V4.0 alignment.

It extracts Course Outcomes (COs) from syllabus PDFs, performs Knowledge Profile (WK1–WK9) mappings, computes GAPC V4.0 Performance Indicator matrix strengths, and serves a multi-user interactive web interface.

---

## 🚀 Key Features

* **AI-Powered Extraction**: Generates 5 precise, measurable Course Outcomes from syllabus PDFs using Gemini AI.
* **Knowledge Profile Mapping**: Maps COs against Washington Accord Knowledge Profiles (WK1–WK9) with detailed keyword justifications.
* **GAPC V4.0 Compliance**: Calculates CO-PO mapping strengths using official formula:
  $$\text{Mapping Strength } X = \left(\frac{\text{Yes PIs}}{\text{Total PIs}}\right) \times 100$$
  * $X = 68\text{--}100\% \rightarrow \mathbf{3}$ (Substantial)
  * $X = 34\text{--}67\% \rightarrow \mathbf{2}$ (Moderate)
  * $X = 10\text{--}33\% \rightarrow \mathbf{1}$ (Slight)
  * $X < 10\% \rightarrow \mathbf{0}$ (No mapping)
* **Web UI Server**: Built-in FastAPI application with Server-Sent Events (SSE) streaming real-time terminal progress directly to the browser.
* **Multi-User Isolation**: Handles concurrent user requests by generating uniquely hashed file names in separate `input/` and `output/` folders.

---

## 📦 Installation

### From GitHub

Install directly using `pip`:

```bash
pip install git+[https://github.com/baburaj-madathil/obe-software.git](https://github.com/baburaj-madathil/obe-software.git)
