"""
CO-WK / CO-PO Mapping utility for NBA / Washington Accord OBE analysis.

Two pipelines:
  1. generate_co_wk_excel  – PDF syllabus → Course Outcomes + WK mapping
  2. generate_co_po_mapping – Excel of COs → CO-PO strength matrix + justifications

Requires: google-genai, pypdf, pandas, openpyxl
"""

from __future__ import annotations

import os
import re
import json
import logging
from typing import Any, Optional

import pandas as pd
from pypdf import PdfReader
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Reference data
# ---------------------------------------------------------------------------

WK_FULL_DATA = [
    {
        "WK": "WK1",
        "Title": "Natural Sciences & Social Sciences",
        "Aspects": "Physics, Chemistry, Biology, Geology, Social Science awareness.",
        "Focus": "Scientific explanation of materials, ecosystems, environment.",
    },
    {
        "WK": "WK2",
        "Title": "Math, Data & Computing",
        "Aspects": "Calculus, Statistics, Numerical analysis, Algorithmic thinking, Programming (Python/MATLAB), Data modelling.",
        "Focus": "Data interpretation, modelling, regression, or trend analysis.",
    },
    {
        "WK": "WK3",
        "Title": "Engineering Fundamentals",
        "Aspects": "Mechanics, Thermodynamics, Fluid mechanics, formulating problems using basic laws.",
        "Focus": "Fundamental laws governing engineering systems.",
    },
    {
        "WK": "WK4",
        "Title": "Discipline-Specific Knowledge",
        "Aspects": "Professional codes, standards, advanced tools, specialist theoretical frameworks.",
        "Focus": "Application of professional standards and codes of practice.",
    },
    {
        "WK": "WK5",
        "Title": "Sustainability & Economics",
        "Aspects": "Environmental impacts, carbon footprint, life-cycle cost, resource reuse, net zero.",
        "Focus": "Sustainable design, green technologies or environmental protection.",
    },
    {
        "WK": "WK6",
        "Title": "Engineering Practice & Technology",
        "Aspects": "Industry tools, materials, lab/fieldwork, practical procedures, hands-on skills.",
        "Focus": "Understanding of processes, materials, equipment, or technologies.",
    },
    {
        "WK": "WK7",
        "Title": "Engineer's Role in Society",
        "Aspects": "Public safety, legal responsibility, societal impact, community needs.",
        "Focus": "Infrastructure for society, disaster resilience.",
    },
    {
        "WK": "WK8",
        "Title": "Research & Innovation",
        "Aspects": "Literature review, critical thinking, creative approaches, emerging issues.",
        "Focus": "Exposure to latest research or trends.",
    },
    {
        "WK": "WK9",
        "Title": "Ethics & Inclusivity",
        "Aspects": "Professional ethics, integrity, diversity, inclusive behavior.",
        "Focus": "Professional conduct and ethical decision-making.",
    },
]

# JSON Schema for structured output (generate_co_wk_excel)
ANALYSIS_SCHEMA = {
    "type": "ARRAY",
    "items": {
        "type": "OBJECT",
        "properties": {
            "CO_Number": {"type": "STRING"},
            "Course_Outcome": {"type": "STRING"},
            "Blooms_Level": {"type": "STRING"},
            "Mapped_WKs": {"type": "ARRAY", "items": {"type": "STRING"}},
            "Primary_WK": {"type": "STRING"},
            "Detailed_Justification": {
                "type": "STRING",
                "description": (
                    "Comprehensive explanation of why these WKs were mapped "
                    "based on the CO text and WK Key Aspects."
                ),
            },
        },
        "required": [
            "CO_Number",
            "Course_Outcome",
            "Blooms_Level",
            "Mapped_WKs",
            "Primary_WK",
            "Detailed_Justification",
        ],
    },
}

# ---------------------------------------------------------------------------
# Configuration defaults
# ---------------------------------------------------------------------------

DEFAULT_MODEL = "gemini-2.5-flash"
DEFAULT_INPUT_EXCEL = "CO_WK_Mapping.xlsx"   # WK pipeline output / PO pipeline input
DEFAULT_OUTPUT_EXCEL = "CO_PO_Mapping.xlsx"  # PO pipeline output

# ---------------------------------------------------------------------------
# GAPC V4.0 – Program Outcomes, Competencies & Performance Indicators
# Source: NBA GAPC V4.0 (Graduate Attributes and Professional Competencies)
#
# Mapping Strength formula (official):
#   X = (Number of Yes) / (Number of PIs) × 100
# Rubrics:
#   X = 10–33  → Mapping Value = 1 (Slight)
#   X = 34–67  → Mapping Value = 2 (Moderate)
#   X = 68–100 → Mapping Value = 3 (Substantial)
#   X < 10 or no Yes → Mapping Value = 0 (No mapping)
# ---------------------------------------------------------------------------

GAPC_PO_CATALOG: dict[str, dict] = {
    "PO1": {
        "title": "Engineering Knowledge",
        "statement": (
            "Apply knowledge of mathematics, natural science, computing, engineering "
            "fundamentals and an engineering specialization as specified in WK1 to WK4 "
            "respectively to develop the solution of complex engineering problems."
        ),
        "wks": ["WK1", "WK2", "WK3", "WK4"],
        "pis": {
            "1.1.1": "Apply mathematical techniques such as calculus, linear algebra, and statistics to solve problems",
            "1.1.2": "Apply advanced mathematical techniques to model and solve engineering problems",
            "1.2.1": "Apply laws of natural science to an engineering problem",
            "1.3.1": "Apply fundamental engineering concepts to solve engineering problems",
            "1.4.1": "Apply specialized engineering knowledge of the program to solve engineering problems",
        },
    },
    "PO2": {
        "title": "Problem Analysis",
        "statement": (
            "Identify, formulate, review research literature and analyze complex engineering "
            "problems reaching substantiated conclusions with consideration for sustainable development."
        ),
        "wks": ["WK1", "WK2", "WK3", "WK4"],
        "pis": {
            "2.1.1": "Articulate problem statements and identify objectives",
            "2.1.2": "Identify engineering systems, variables, and parameters to solve the problems",
            "2.1.3": "Identify the mathematical, engineering and other relevant knowledge that applies to a given problem",
            "2.2.1": "Reframe complex problems into interconnected sub-problems",
            "2.2.2": "Identify, assemble and evaluate information and resources",
            "2.2.3": "Identify existing processes/solution methods for solving the problem, including forming justified approximations and assumptions",
            "2.2.4": "Compare and contrast alternative solution processes to select the best process that can also satisfy the technical, socio-economic and environmental dimensions of sustainability",
            "2.3.1": "Combine scientific principles and engineering concepts to formulate model/s of a system or process appropriate in terms of applicability and required accuracy",
            "2.3.2": "Identify assumptions (mathematical and physical) necessary to allow modeling of a system at the level of accuracy required",
            "2.4.1": "Apply engineering mathematics and computations to solve mathematical models",
            "2.4.2": "Produce and validate results through skillful use of contemporary engineering tools and models",
            "2.4.3": "Identify sources of error in the solution process, and limitations of the solution",
            "2.4.4": "Extract desired understanding and conclusions consistent with objectives and limitations of the analysis",
        },
    },
    "PO3": {
        "title": "Design/Development of Solutions",
        "statement": (
            "Design creative solutions for complex engineering problems and design/develop "
            "systems/components/processes to meet identified needs with consideration for public "
            "health and safety, whole-life cost, net zero carbon, culture, society and environment."
        ),
        "wks": ["WK5"],
        "pis": {
            "3.1.1": "Recognize that need analysis is key to good problem definition",
            "3.1.2": "Elicit and document engineering requirements from stakeholders",
            "3.1.3": "Synthesize engineering requirements from a review of the state-of-the-art",
            "3.1.4": "Extract engineering requirements from relevant engineering Codes and Standards",
            "3.1.5": "Explore and synthesize engineering requirements considering health, safety risks, environmental, cultural and societal issues",
            "3.1.6": "Determine design objectives, functional requirements and arrive at specifications",
            "3.2.1": "Apply formal idea generation tools to develop multiple engineering design solutions",
            "3.2.2": "Build models/prototypes to develop a diverse set of design solutions",
            "3.2.3": "Identify suitable criteria for the evaluation of alternate design solutions",
            "3.3.1": "Apply formal decision-making tools to select optimal engineering design solutions for further development",
            "3.3.2": "Consult with domain experts and stakeholders to select candidate engineering design solution",
            "3.4.1": "Refine a conceptual design into a detailed design within existing constraints with consideration for public health and safety, whole-life cost, net-zero carbon, culture, society and environment",
        },
    },
    "PO4": {
        "title": "Conduct Investigations of Complex Problems",
        "statement": (
            "Conduct investigations of complex engineering problems using research-based knowledge "
            "including design of experiments, modelling, analysis and interpretation of data to provide valid conclusions."
        ),
        "wks": ["WK8"],
        "pis": {
            "4.1.1": "Define a problem, its scope and importance for purposes of investigation",
            "4.1.2": "Examine the relevant methods, tools and techniques of experiment design, system calibration, data acquisition, analysis and presentation",
            "4.1.3": "Apply appropriate instrumentation and/or software tools to make measurements of physical quantities",
            "4.1.4": "Establish a relationship between measured data and underlying physical principles",
            "4.2.1": "Design and develop an experimental approach, specify appropriate equipment and procedures",
            "4.2.2": "Understand the importance of the statistical design of experiments and choose an appropriate experimental design plan",
            "4.3.1": "Use appropriate procedures, tools and techniques to conduct experiments and collect data",
            "4.3.2": "Analyze data for trends and correlations, stating possible errors and limitations",
            "4.3.3": "Represent data (tabular and/or graphical) so as to facilitate analysis, explanation and drawing of conclusions",
        },
    },
    "PO5": {
        "title": "Engineering Tool Usage",
        "statement": (
            "Create, select and apply appropriate techniques, resources and modern engineering and IT tools, "
            "including prediction and modelling, recognizing their limitations to solve complex engineering problems."
        ),
        "wks": ["WK2", "WK6"],
        "pis": {
            "5.1.1": "Identify modern engineering tools such as computer-aided drafting, modeling and analysis; techniques and resources for engineering activities",
            "5.1.2": "Create/adapt/modify/extend tools and techniques to solve engineering problems",
            "5.2.1": "Identify the strengths and limitations of tools for acquiring information, modeling and simulating, monitoring system performance, and creating engineering designs",
            "5.2.2": "Demonstrate proficiency in using discipline-specific tools",
            "5.3.1": "Discuss limitations and validate tools, techniques and resources",
            "5.3.2": "Verify the credibility of results from tool use with reference to accuracy, limitations, and inherent assumptions",
        },
    },
    "PO6": {
        "title": "The Engineer and The World",
        "statement": (
            "Analyze and evaluate societal and environmental aspects while solving complex engineering problems "
            "for its impact on sustainability with reference to economy, health, safety, legal framework, culture and environment."
        ),
        "wks": ["WK1", "WK5", "WK7"],
        "pis": {
            "6.1.1": "Identify and describe various engineering roles; particularly for impact on sustainability with reference to economy, health, safety, legal framework, culture and environment",
            "6.2.1": "Interpret legislation, regulations, codes, and standards relevant to the discipline and explain contribution to protection of the public",
            "6.3.1": "Identify risks/impacts in the life-cycle of an engineering product or activity",
            "6.3.2": "Understand the relationship between the technical, socio-economic and environmental dimensions of sustainability",
        },
    },
    "PO7": {
        "title": "Ethics",
        "statement": (
            "Apply ethical principles and commit to professional ethics, human values, diversity and inclusion; "
            "adhere to national and international laws."
        ),
        "wks": ["WK9"],
        "pis": {
            "7.1.1": "Identify situations of unethical professional conduct and propose ethical alternatives",
            "7.1.2": "Understand the need for diversity by reason of ethnicity, gender, age, physical ability etc. with mutual understanding, respect and inclusive attitudes",
            "7.2.1": "Identify tenets of the professional code of ethics",
            "7.2.2": "Examine and apply moral and ethical principles to known case studies",
        },
    },
    "PO8": {
        "title": "Individual and Collaborative Team Work",
        "statement": "Function effectively as an individual, and as a member or leader in diverse/multi-disciplinary teams.",
        "wks": [],
        "pis": {
            "8.1.1": "Recognize a variety of working and learning preferences; appreciate the value of diversity on a team",
            "8.1.2": "Implement the norms of practice (rules, roles, charters, agendas) of effective team work to accomplish a goal",
            "8.2.1": "Demonstrate effective communication, problem-solving, conflict resolution and leadership skills",
            "8.2.2": "Treat other team members respectfully",
            "8.2.3": "Listen to other members",
            "8.2.4": "Maintain composure in difficult situations",
            "8.3.1": "Present results as a team, with smooth integration of contributions from all individual efforts",
        },
    },
    "PO9": {
        "title": "Communication",
        "statement": (
            "Communicate effectively and inclusively within the engineering community and society at large, "
            "such as being able to comprehend and write effective reports and design documentation, make effective "
            "presentations considering cultural, language, and learning differences."
        ),
        "wks": [],
        "pis": {
            "9.1.1": "Read, understand and interpret technical and non-technical information",
            "9.1.2": "Produce clear, well-constructed, and well-supported written engineering documents",
            "9.1.3": "Create flow in a document or presentation – a logical progression of ideas so that the main point is clear",
            "9.2.1": "Listen to and comprehend information, instructions, and viewpoints of others",
            "9.2.2": "Deliver effective oral presentations to technical and non-technical audiences",
            "9.3.1": "Create engineering-standard figures, reports and drawings to complement writing and presentations",
            "9.3.2": "Use a variety of media effectively to convey a message in a document or a presentation",
        },
    },
    "PO10": {
        "title": "Project Management and Finance",
        "statement": (
            "Apply knowledge and understanding of engineering management principles and economic decision-making "
            "and apply these to one's own work, as a member and leader in a team, and to manage projects and in multidisciplinary environments."
        ),
        "wks": [],
        "pis": {
            "10.1.1": "Describe various economic and financial costs/benefits of an engineering activity",
            "10.1.2": "Analyze different forms of financial statements to evaluate the financial status of an engineering project",
            "10.2.1": "Analyze and select the most appropriate proposal based on economic and financial considerations",
            "10.3.1": "Identify the tasks required to complete an engineering activity, and the resources required to complete the tasks",
            "10.3.2": "Use project management tools to schedule an engineering project so it is completed on time and on budget",
        },
    },
    "PO11": {
        "title": "Life Long Learning",
        "statement": (
            "Recognize the need for and have the preparation and ability for (i) independent and life-long learning "
            "(ii) adaptability to new and emerging technologies and (iii) critical thinking in the broadest context of technological change."
        ),
        "wks": ["WK8"],
        "pis": {
            "11.1.1": "Describe the rationale for the requirement for continuing professional development",
            "11.1.2": "Identify deficiencies or gaps in knowledge and demonstrate an ability to source information to close this gap",
            "11.2.1": "Identify historic points of technological advance in engineering that required practitioners to seek education to stay updated",
            "11.2.2": "Recognize the need and be able to clearly explain why it is vitally important to keep current regarding new developments in your field",
            "11.3.1": "Source and comprehend technical literature and other credible sources of information",
            "11.3.2": "Analyze sourced technical and popular information for feasibility, viability, sustainability, etc.",
        },
    },
}


def gapc_mapping_value(yes_count: int, total_pis: int) -> tuple[int, float]:
    """
    GAPC V4.0 official strength calculation.

    X = (Number of Yes) / (Number of PIs) × 100
    Returns (mapping_value, X_percent).
      mapping_value: 0, 1, 2, or 3
    """
    if total_pis <= 0 or yes_count <= 0:
        return 0, 0.0
    x = (yes_count / total_pis) * 100.0
    if x < 10:
        return 0, round(x, 1)
    if x <= 33:
        return 1, round(x, 1)
    if x <= 67:
        return 2, round(x, 1)
    return 3, round(x, 1)

# ---------------------------------------------------------------------------
# Styling (module-level for convenience in a script; safe for single-threaded use)
# ---------------------------------------------------------------------------

header_font = Font(bold=True, color="FFFFFF", size=11)
header_fill = PatternFill("solid", fgColor="1F4E79")
s3_fill = PatternFill("solid", fgColor="63BE7B")
s2_fill = PatternFill("solid", fgColor="FFEB84")
s1_fill = PatternFill("solid", fgColor="F4B183")
no_fill = PatternFill("solid", fgColor="D9D9D9")
error_fill = PatternFill("solid", fgColor="FF6B6B")
thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _style_header(ws, row: int, start_col: int, end_col: int) -> None:
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _resolve_api_key(api_key: str) -> str:
    """Validate and return the API key. api_key is mandatory."""
    if not api_key or not str(api_key).strip() or str(api_key).strip() in ("", "YOUR_GEMINI_API_KEY_HERE"):
        raise ValueError(
            "api_key is required. Pass a valid Gemini API key "
            "(e.g. from https://aistudio.google.com/apikey)."
        )
    return str(api_key).strip()


def _strip_markdown_fences(text: str) -> str:
    """Remove common markdown code-fence wrappers from model output."""
    text = text.strip()
    # ```json ... ``` or ``` ... ```
    match = re.match(r"^```(?:json)?\s*\n?(.*?)\n?```\s*$", text, re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1).strip()
    # Fallback: strip leading/trailing fences that may be incomplete
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, count=1, flags=re.IGNORECASE)
    if text.endswith("```"):
        text = text[: -3]
    return text.strip()


def _po_sort_key(po_name: str) -> int:
    """Safe numeric sort key for keys like 'PO1', 'PO12'. Unknown keys go last."""
    if len(po_name) > 2 and po_name[2:].isdigit():
        return int(po_name[2:])
    return 999


# ---------------------------------------------------------------------------
# PDF extraction
# ---------------------------------------------------------------------------

def extract_syllabus_text(pdf_path: str) -> str:
    """Extract text content from a PDF file."""
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Source PDF file not found at: {pdf_path}")

    reader = PdfReader(pdf_path)
    text_blocks = [p.extract_text() for p in reader.pages if p.extract_text()]
    return "\n".join(text_blocks)


# ---------------------------------------------------------------------------
# Pipeline 1: PDF → CO + WK mapping
# ---------------------------------------------------------------------------

def generate_co_wk_excel(
    pdf_path: str,
    model_name: str,
    api_key: str,
    output_excel_path: str = DEFAULT_INPUT_EXCEL,  # CO_WK_Mapping.xlsx
) -> pd.DataFrame:
    """
    Process a syllabus PDF with Gemini: extract/generate Course Outcomes (CO1–CO5)
    and map them to Washington Accord Knowledge Profiles (WK1–WK9).
    Writes an Excel file and returns the DataFrame.

    Parameters
    ----------
    pdf_path : str
        Path to the syllabus PDF.
    model_name : str
        **Required.** Gemini model id (e.g. "gemini-2.5-flash").
    api_key : str
        **Required.** Gemini API key.
    output_excel_path : str
        Destination Excel path.
    """
    if not model_name or not str(model_name).strip():
        raise ValueError("model_name is required (e.g. 'gemini-2.5-flash').")

    if genai is None:
        raise ImportError(
            "Required dependency 'google-genai' is missing. "
            "Install via: pip install google-genai pypdf pandas openpyxl"
        )

    effective_api_key = _resolve_api_key(api_key)

    print(f"[Step 1/5] Extracting content from PDF: {pdf_path}")
    syllabus_text = extract_syllabus_text(pdf_path)
    print(f"[Step 1/5] Extraction complete ({len(syllabus_text)} characters extracted).")

    print(f"[Step 2/5] Initializing Gemini client using model: {model_name}...")
    client = genai.Client(api_key=effective_api_key)

    prompt = (
        "You are an expert Accreditation Auditor. Perform a dual task:\n"
        "1. Generate 5 precise, measurable Course Outcomes (CO1-CO5) based on the syllabus provided.\n"
        "2. For each CO, map EVERY possible Washington Accord Knowledge Profile (WK1-WK9) that applies.\n"
        "3. Provide a 'Detailed_Justification' for each CO. You MUST explicitly mention which keywords in the CO "
        "relate to the 'Key Aspects' or 'Focus' of the mapped WKs.\n\n"
        f"WK DEFINITIONS:\n{json.dumps(WK_FULL_DATA, indent=2)}\n\n"
        f"SYLLABUS CONTENT:\n{syllabus_text}"
    )

    print("[Step 3/5] Requesting CO extraction and WK mapping from Gemini...")
    response = client.models.generate_content(
        model=model_name,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ANALYSIS_SCHEMA,
        ),
    )
    print("[Step 3/5] AI analysis finished and validated against schema.")

    print("[Step 4/5] Constructing tabular DataFrame structure...")
    raw_data = response.parsed
    if not raw_data or not isinstance(raw_data, (list, tuple)):
        raise ValueError(
            "Gemini returned empty or invalid structured output. "
            "Check model name, schema compatibility, and API response."
        )

    table_rows = []
    for item in raw_data:
        # Support both attribute-style and dict-style parsed items
        if isinstance(item, dict):
            get = item.get
        else:
            get = lambda k, default="": getattr(item, k, default)

        mapped_wks = get("Mapped_WKs", []) or []
        if not isinstance(mapped_wks, (list, tuple)):
            mapped_wks = [str(mapped_wks)]

        table_rows.append(
            {
                "CO ID": get("CO_Number", ""),
                "Course Outcome": get("Course_Outcome", ""),
                "Bloom's Level": get("Blooms_Level", ""),
                "Mapped WKs": ", ".join(str(w) for w in mapped_wks),
                "Primary WK": get("Primary_WK", ""),
                "Justification": get("Detailed_Justification", ""),
            }
        )

    df = pd.DataFrame(table_rows)

    print(f"[Step 5/5] Exporting mapping results to Excel file: {output_excel_path}")
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="CO_WK_Mapping", index=False)

    csv_path = os.path.splitext(output_excel_path)[0] + ".csv"
    df.to_csv(csv_path, index=False)

    print(f"✅ SUCCESS: OBE analysis finished. Excel generated at: {output_excel_path}")
    return df


# ---------------------------------------------------------------------------
# Pipeline 2 helpers: Gemini client + single-CO analysis
# ---------------------------------------------------------------------------

def init_gemini_client(api_key: str):
    """Initialize and return the Google GenAI client. api_key is mandatory."""
    if genai is None:
        raise ImportError(
            "Required dependency 'google-genai' is missing. "
            "Install via: pip install google-genai"
        )
    key = _resolve_api_key(api_key)
    return genai.Client(api_key=key)


def analyze_co_with_gemini(
    client,
    co_id: str,
    statement: str,
    bloom: str,
    existing_wks: str,
    model_name: str,
) -> dict:
    """
    Send one Course Outcome to Gemini and return a structured mapping dict.

    Strength is computed using the official GAPC V4.0 formula:
      X = (Number of Yes PIs) / (Total PIs for that PO) × 100
      X 10-33 → 1 | 34-67 → 2 | 68-100 → 3 | else → 0

    On failure returns {"co_id": ..., "error": "..."}.
    """
    # Build compact PI catalog for the prompt
    pi_catalog_text = []
    for po, info in GAPC_PO_CATALOG.items():
        pi_lines = [f"  {pid}: {desc}" for pid, desc in info["pis"].items()]
        pi_catalog_text.append(
            f"{po} ({info['title']}) – {len(info['pis'])} PIs:\n" + "\n".join(pi_lines)
        )
    catalog_block = "\n\n".join(pi_catalog_text)

    system_prompt = f"""
You are an expert in NBA Outcome-Based Education and GAPC V4.0.

Your task is to evaluate ONE Course Outcome against every Program Outcome (PO1–PO11)
by marking each Performance Indicator (PI) as Yes or No.

OFFICIAL STRENGTH FORMULA (do NOT invent strengths):
  X = (Number of Yes) / (Number of PIs for that PO) × 100
  Rubrics: X=10-33 → mapping_value=1; X=34-67 → 2; X=68-100 → 3; else 0.

For each PO that the CO genuinely addresses, return the list of PI IDs that are Yes.
Only mark a PI as Yes when the CO statement clearly supports that measurable action.
Do NOT mark PIs as Yes just because the PO title sounds related.

Return ONLY a valid JSON object with this exact structure:

{{
  "co_id": "CO1",
  "understanding": "Clear explanation of what the student is expected to do",
  "suggested_wks": ["WK2", "WK3", "WK4"],
  "primary_wk": "WK3",
  "pi_yes": {{
    "PO1": ["1.1.1", "1.3.1"],
    "PO2": ["2.1.1", "2.1.2", "2.4.1"]
  }},
  "justifications": {{
    "PO1": "Why these PIs are supported by the CO statement...",
    "PO2": "..."
  }},
  "recommendations": "Suggestions to improve the CO statement if needed"
}}

Rules:
- pi_yes must only contain POs that have at least one Yes PI
- PI IDs must exactly match the catalog below
- Return pure JSON only (no markdown)

GAPC V4.0 PERFORMANCE INDICATOR CATALOG:
{catalog_block}
"""

    prompt = f"""
{system_prompt}

CO ID: {co_id}
Statement: {statement}
Bloom's Level: {bloom}
Existing WK mapping: {existing_wks}

Return ONLY valid JSON with pi_yes and justifications.
"""
    try:
        response = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.2,
                response_mime_type="application/json",
            ),
        )
        text = _strip_markdown_fences(response.text or "")
        raw = json.loads(text)

        # Post-process: compute official GAPC strength from PI Yes lists
        pi_yes = raw.get("pi_yes") or {}
        justifications = raw.get("justifications") or {}
        mapped_pos: dict[str, dict] = {}

        for po, yes_list in pi_yes.items():
            if po not in GAPC_PO_CATALOG:
                continue
            catalog_pis = GAPC_PO_CATALOG[po]["pis"]
            # Keep only valid PI IDs
            valid_yes = [pid for pid in (yes_list or []) if pid in catalog_pis]
            total = len(catalog_pis)
            yes_count = len(valid_yes)
            strength, x_pct = gapc_mapping_value(yes_count, total)
            if strength == 0:
                continue
            mapped_pos[po] = {
                "strength": strength,
                "mapping_strength_pct": x_pct,
                "yes_count": yes_count,
                "total_pis": total,
                "contributing_pis": valid_yes,
                "justification": justifications.get(po, ""),
            }

        return {
            "co_id": raw.get("co_id", co_id),
            "understanding": raw.get("understanding", ""),
            "suggested_wks": raw.get("suggested_wks") or [],
            "primary_wk": raw.get("primary_wk", ""),
            "mapped_pos": mapped_pos,
            "recommendations": raw.get("recommendations", ""),
        }
    except Exception as e:
        logger.exception("Error analyzing %s", co_id)
        print(f"Error analyzing {co_id}: {e}")
        return {"co_id": co_id, "error": str(e)}


# ---------------------------------------------------------------------------
# Read COs from Excel
# ---------------------------------------------------------------------------

def read_cos_from_excel(excel_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """
    Read Course Outcomes from an Excel file.
    Expected columns (flexible matching): CO ID / CO, Course Outcome / Statement,
    Bloom's Level, Mapped WKs.
    """
    if sheet_name:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    else:
        df = pd.read_excel(excel_path)

    col_map = {}
    for col in df.columns:
        col_lower = str(col).strip().lower()
        if "co id" in col_lower or col_lower == "co":
            col_map[col] = "CO ID"
        elif (
            "course outcome" in col_lower
            or "co statement" in col_lower
            or col_lower == "statement"
        ):
            col_map[col] = "Course Outcome"
        elif "bloom" in col_lower:
            col_map[col] = "Bloom's Level"
        # Narrow match – avoid catching "Primary WK", "Weekly …", etc.
        elif "mapped wk" in col_lower or col_lower in ("mapped_wks", "wks", "mapped wks"):
            col_map[col] = "Mapped WKs"

    df = df.rename(columns=col_map)

    required = ["CO ID", "Course Outcome"]
    for col in required:
        if col not in df.columns:
            raise ValueError(
                f"Required column '{col}' not found in the Excel file. "
                f"Available columns: {list(df.columns)}"
            )

    return df


# ---------------------------------------------------------------------------
# Analyze all COs
# ---------------------------------------------------------------------------

def analyze_all_cos(
    excel_path: str,
    client,
    model_name: str,
    sheet_name: Optional[str] = None,
) -> list[dict]:
    """Read Excel and analyze every CO using Gemini. model_name is required."""
    if not model_name or not str(model_name).strip():
        raise ValueError("model_name is required (e.g. 'gemini-2.5-flash').")

    df = read_cos_from_excel(excel_path, sheet_name)
    results: list[dict] = []

    print(f"\nFound {len(df)} Course Outcomes. Sending to Gemini ({model_name})...\n")

    for _, row in df.iterrows():
        co_id = str(row["CO ID"]).strip()
        statement = str(row["Course Outcome"]).strip()
        bloom = str(row.get("Bloom's Level", "") or "").strip()
        existing_wks = str(row.get("Mapped WKs", "") or "").strip()

        print(f"  Analyzing {co_id} ...")
        result = analyze_co_with_gemini(
            client, co_id, statement, bloom, existing_wks, model_name=model_name
        )
        results.append(result)
        status = "ERROR" if "error" in result else "Done"
        print(f"  → {status}")

    return results


# ---------------------------------------------------------------------------
# Excel report generation
# ---------------------------------------------------------------------------

def _compute_strength_stats(results: list[dict], all_pos: list[str]) -> dict:
    """
    Compute CO-PO strength statistics used by the Strength Calculation sheet.

    Returns a dict with:
      - co_stats: list of per-CO metrics
      - po_stats: dict keyed by PO with count_3/2/1, total_mapped, sum_strength, avg_strength
      - overall: aggregate totals
    """
    co_stats = []
    po_stats = {
        po: {"count_3": 0, "count_2": 0, "count_1": 0, "total_mapped": 0, "sum_strength": 0}
        for po in all_pos
    }
    total_mappings = 0
    sum_all_strengths = 0
    errors = 0

    for res in results:
        co = res.get("co_id", "")
        if "error" in res:
            errors += 1
            co_stats.append(
                {
                    "co_id": co,
                    "mapped_count": 0,
                    "sum_strength": 0,
                    "avg_strength": None,
                    "count_3": 0,
                    "count_2": 0,
                    "count_1": 0,
                    "error": True,
                }
            )
            continue

        mapped = res.get("mapped_pos") or {}
        c3 = c2 = c1 = 0
        ssum = 0
        for po in all_pos:
            data = mapped.get(po) or {}
            strength = data.get("strength")
            if strength not in (1, 2, 3):
                continue
            ssum += strength
            total_mappings += 1
            sum_all_strengths += strength
            po_stats[po]["total_mapped"] += 1
            po_stats[po]["sum_strength"] += strength
            if strength == 3:
                c3 += 1
                po_stats[po]["count_3"] += 1
            elif strength == 2:
                c2 += 1
                po_stats[po]["count_2"] += 1
            else:
                c1 += 1
                po_stats[po]["count_1"] += 1

        mapped_count = c3 + c2 + c1
        co_stats.append(
            {
                "co_id": co,
                "mapped_count": mapped_count,
                "sum_strength": ssum,
                "avg_strength": round(ssum / mapped_count, 2) if mapped_count else None,
                "count_3": c3,
                "count_2": c2,
                "count_1": c1,
                "error": False,
            }
        )

    for po, st in po_stats.items():
        st["avg_strength"] = (
            round(st["sum_strength"] / st["total_mapped"], 2) if st["total_mapped"] else None
        )

    overall = {
        "total_cos": len(results),
        "errors": errors,
        "total_mappings": total_mappings,
        "sum_all_strengths": sum_all_strengths,
        "avg_strength": (
            round(sum_all_strengths / total_mappings, 2) if total_mappings else None
        ),
        "pos_covered": sum(1 for st in po_stats.values() if st["total_mapped"] > 0),
    }
    return {"co_stats": co_stats, "po_stats": po_stats, "overall": overall}


def create_excel_report(
    results: list[dict],
    model_name: str,
    output_path: str = DEFAULT_OUTPUT_EXCEL,
    input_excel: Optional[str] = None,
) -> None:
    """Create a multi-sheet Excel file with strength matrix, calculations, justifications, etc.

    If input_excel is provided and exists, its content is copied into the first
    sheet (0_Source_CO_WK) so the output workbook is self-contained.
    """
    if not model_name or not str(model_name).strip():
        raise ValueError("model_name is required.")
    wb = Workbook()
    ALL_POS = [f"PO{i}" for i in range(1, 12)]
    stats = _compute_strength_stats(results, ALL_POS)

    # ----- Sheet 0: Source CO-WK data from input Excel (if provided) -----
    if input_excel and os.path.exists(input_excel):
        try:
            src_df = pd.read_excel(input_excel)
            ws0 = wb.active
            ws0.title = "0_Source_CO_WK"
            ws0["A1"] = f"Source CO-WK Mapping (from {os.path.basename(input_excel)})"
            ws0["A1"].font = Font(bold=True, size=13, color="1F4E79")
            ws0.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(src_df.columns), 1))

            # Header row
            for col_idx, col_name in enumerate(src_df.columns, 1):
                cell = ws0.cell(row=3, column=col_idx, value=str(col_name))
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center
                cell.border = thin_border

            # Data rows
            for r_idx, row in enumerate(src_df.itertuples(index=False), 4):
                for c_idx, value in enumerate(row, 1):
                    cell = ws0.cell(row=r_idx, column=c_idx, value=value if pd.notna(value) else "")
                    cell.border = thin_border
                    cell.alignment = left_align if c_idx > 1 else center
                ws0.row_dimensions[r_idx].height = 40

            for c_idx in range(1, len(src_df.columns) + 1):
                ws0.column_dimensions[get_column_letter(c_idx)].width = 18
            # Wider columns for text-heavy fields
            for c_idx, col_name in enumerate(src_df.columns, 1):
                name_l = str(col_name).lower()
                if "outcome" in name_l or "justif" in name_l:
                    ws0.column_dimensions[get_column_letter(c_idx)].width = 55
                elif "mapped" in name_l or "wk" in name_l:
                    ws0.column_dimensions[get_column_letter(c_idx)].width = 22

            # Create strength matrix as next sheet
            ws1 = wb.create_sheet("1_Strength_Matrix")
        except Exception as e:
            logger.warning("Could not embed input Excel %s: %s", input_excel, e)
            ws1 = wb.active
            ws1.title = "1_Strength_Matrix"
    else:
        ws1 = wb.active
        ws1.title = "1_Strength_Matrix"

    ws1["A1"] = "CO-PO Correlation Strength Matrix (GAPC V4.0)"
    ws1["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws1.merge_cells("A1:L1")

    ws1["A2"] = (
        "3 = Substantial (X=68–100%) | 2 = Moderate (X=34–67%) | 1 = Slight (X=10–33%) | – = No correlation  "
        "where X = (Yes PIs / Total PIs) × 100"
    )
    ws1["A2"].font = Font(italic=True, size=9)

    ws1.cell(row=4, column=1, value="CO")
    for j, po in enumerate(ALL_POS, 2):
        ws1.cell(row=4, column=j, value=po)
    _style_header(ws1, 4, 1, 12)

    for i, res in enumerate(results):
        r = 5 + i
        co = res.get("co_id", f"CO{i + 1}")
        cell_co = ws1.cell(row=r, column=1, value=co)
        cell_co.font = Font(bold=True)
        cell_co.alignment = center
        cell_co.border = thin_border

        if "error" in res:
            for j in range(2, 13):
                cell = ws1.cell(row=r, column=j, value="ERR")
                cell.fill = error_fill
                cell.alignment = center
                cell.border = thin_border
                cell.font = Font(bold=True)
            continue

        mapped = res.get("mapped_pos") or {}
        for j, po in enumerate(ALL_POS, 2):
            cell = ws1.cell(row=r, column=j)
            strength = (mapped.get(po) or {}).get("strength")
            if strength == 3:
                cell.value = 3
                cell.fill = s3_fill
            elif strength == 2:
                cell.value = 2
                cell.fill = s2_fill
            elif strength == 1:
                cell.value = 1
                cell.fill = s1_fill
            else:
                cell.value = "–"
                cell.fill = no_fill
            cell.alignment = center
            cell.border = thin_border
            cell.font = Font(bold=True)

    ws1.column_dimensions["A"].width = 10
    for col in range(2, 13):
        ws1.column_dimensions[get_column_letter(col)].width = 7

    # ----- Sheet 2: Strength Calculation (NEW) -----
    ws_calc = wb.create_sheet("2_Strength_Calculation")

    ws_calc["A1"] = "CO-PO Mapping Strength Calculation (GAPC V4.0)"
    ws_calc["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws_calc.merge_cells("A1:H1")

    ws_calc["A2"] = (
        "Official GAPC V4.0 formula: X = (Number of Yes PIs) / (Total PIs for that PO) × 100.  "
        "Rubrics: X=10–33 → 1 (Slight) | X=34–67 → 2 (Moderate) | X=68–100 → 3 (Substantial) | X<10 → 0 (No mapping).  "
        "Average Strength = Sum of mapping values / Number of mapped POs (unmapped cells excluded)."
    )
    ws_calc["A2"].font = Font(italic=True, size=9)
    ws_calc.merge_cells("A2:H2")

    # --- Section A: Per-CO calculation ---
    ws_calc["A4"] = "A. Per-CO Strength Calculation"
    ws_calc["A4"].font = Font(bold=True, size=11, color="1F4E79")

    co_headers = [
        "CO",
        "POs Mapped",
        "Count (3)",
        "Count (2)",
        "Count (1)",
        "Sum of Strengths",
        "Average Strength",
        "Status",
    ]
    for col, h in enumerate(co_headers, 1):
        ws_calc.cell(row=5, column=col, value=h)
    _style_header(ws_calc, 5, 1, 8)

    for i, cs in enumerate(stats["co_stats"]):
        r = 6 + i
        ws_calc.cell(row=r, column=1, value=cs["co_id"]).alignment = center
        ws_calc.cell(row=r, column=2, value=cs["mapped_count"]).alignment = center
        ws_calc.cell(row=r, column=3, value=cs["count_3"]).alignment = center
        ws_calc.cell(row=r, column=4, value=cs["count_2"]).alignment = center
        ws_calc.cell(row=r, column=5, value=cs["count_1"]).alignment = center
        ws_calc.cell(row=r, column=6, value=cs["sum_strength"]).alignment = center
        avg_cell = ws_calc.cell(
            row=r,
            column=7,
            value=cs["avg_strength"] if cs["avg_strength"] is not None else "–",
        )
        avg_cell.alignment = center
        if cs.get("error"):
            status_cell = ws_calc.cell(row=r, column=8, value="ERROR")
            status_cell.fill = error_fill
        else:
            status_cell = ws_calc.cell(row=r, column=8, value="OK")
            if cs["avg_strength"] is not None and cs["avg_strength"] >= 2.5:
                avg_cell.fill = s3_fill
            elif cs["avg_strength"] is not None and cs["avg_strength"] >= 1.5:
                avg_cell.fill = s2_fill
            elif cs["avg_strength"] is not None:
                avg_cell.fill = s1_fill
        status_cell.alignment = center
        for c in range(1, 9):
            ws_calc.cell(row=r, column=c).border = thin_border

    # --- Section B: Per-PO calculation ---
    po_start = 6 + len(stats["co_stats"]) + 2
    ws_calc.cell(row=po_start, column=1, value="B. Per-PO Strength Calculation")
    ws_calc.cell(row=po_start, column=1).font = Font(bold=True, size=11, color="1F4E79")

    po_headers = [
        "PO",
        "COs Mapped",
        "Count (3)",
        "Count (2)",
        "Count (1)",
        "Sum of Strengths",
        "Average Strength",
        "Coverage",
    ]
    header_row = po_start + 1
    for col, h in enumerate(po_headers, 1):
        ws_calc.cell(row=header_row, column=col, value=h)
    _style_header(ws_calc, header_row, 1, 8)

    for i, po in enumerate(ALL_POS):
        r = header_row + 1 + i
        st = stats["po_stats"][po]
        ws_calc.cell(row=r, column=1, value=po).alignment = center
        ws_calc.cell(row=r, column=2, value=st["total_mapped"]).alignment = center
        ws_calc.cell(row=r, column=3, value=st["count_3"]).alignment = center
        ws_calc.cell(row=r, column=4, value=st["count_2"]).alignment = center
        ws_calc.cell(row=r, column=5, value=st["count_1"]).alignment = center
        ws_calc.cell(row=r, column=6, value=st["sum_strength"]).alignment = center
        avg = st["avg_strength"]
        avg_cell = ws_calc.cell(row=r, column=7, value=avg if avg is not None else "–")
        avg_cell.alignment = center
        if avg is not None:
            if avg >= 2.5:
                avg_cell.fill = s3_fill
            elif avg >= 1.5:
                avg_cell.fill = s2_fill
            else:
                avg_cell.fill = s1_fill
        coverage = "Yes" if st["total_mapped"] > 0 else "No"
        cov_cell = ws_calc.cell(row=r, column=8, value=coverage)
        cov_cell.alignment = center
        if coverage == "No":
            cov_cell.fill = no_fill
        for c in range(1, 9):
            ws_calc.cell(row=r, column=c).border = thin_border

    # --- Section C: Overall summary metrics ---
    overall_start = header_row + 1 + len(ALL_POS) + 2
    ws_calc.cell(row=overall_start, column=1, value="C. Overall Strength Metrics")
    ws_calc.cell(row=overall_start, column=1).font = Font(bold=True, size=11, color="1F4E79")

    ov = stats["overall"]
    metrics = [
        ("Total COs analyzed", ov["total_cos"]),
        ("COs with errors", ov["errors"]),
        ("Total CO-PO mappings", ov["total_mappings"]),
        ("Sum of all strengths", ov["sum_all_strengths"]),
        ("Overall average strength", ov["avg_strength"] if ov["avg_strength"] is not None else "–"),
        ("POs covered (of 11)", f"{ov['pos_covered']} / 11"),
    ]
    for i, (label, value) in enumerate(metrics):
        r = overall_start + 1 + i
        ws_calc.cell(row=r, column=1, value=label).border = thin_border
        val_cell = ws_calc.cell(row=r, column=2, value=value)
        val_cell.alignment = center
        val_cell.border = thin_border

    ws_calc.column_dimensions["A"].width = 22
    ws_calc.column_dimensions["B"].width = 14
    ws_calc.column_dimensions["C"].width = 12
    ws_calc.column_dimensions["D"].width = 12
    ws_calc.column_dimensions["E"].width = 12
    ws_calc.column_dimensions["F"].width = 16
    ws_calc.column_dimensions["G"].width = 16
    ws_calc.column_dimensions["H"].width = 12

    # ----- Sheet 3: Mapping with Justification -----
    ws2 = wb.create_sheet("3_Mapping_with_Justification")

    ws2["A1"] = "Complete CO-PO Mapping with Justification"
    ws2["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws2.merge_cells("A1:F1")

    headers2 = [
        "CO", "PO", "Strength", "X (%)", "Yes/Total PIs",
        "Contributing PIs", "Justification", "Primary WK",
    ]
    for col, h in enumerate(headers2, 1):
        ws2.cell(row=3, column=col, value=h)
    _style_header(ws2, 3, 1, 8)

    row_num = 4
    for res in results:
        co = res.get("co_id", "")
        primary_wk = res.get("primary_wk", "")

        if "error" in res:
            ws2.cell(row=row_num, column=1, value=co).alignment = center
            ws2.cell(row=row_num, column=2, value="—").alignment = center
            ws2.cell(row=row_num, column=3, value="ERROR").alignment = center
            ws2.cell(row=row_num, column=3).fill = error_fill
            for c in range(4, 8):
                ws2.cell(row=row_num, column=c, value="").alignment = center
            ws2.cell(row=row_num, column=7, value=res.get("error", "Unknown error")).alignment = left_align
            for c in range(1, 9):
                ws2.cell(row=row_num, column=c).border = thin_border
            ws2.row_dimensions[row_num].height = 40
            row_num += 1
            continue

        mapped = res.get("mapped_pos") or {}
        if not mapped:
            ws2.cell(row=row_num, column=1, value=co).alignment = center
            ws2.cell(row=row_num, column=2, value="—").alignment = center
            ws2.cell(row=row_num, column=3, value="–").alignment = center
            for c in range(4, 8):
                ws2.cell(row=row_num, column=c, value="").alignment = center
            ws2.cell(row=row_num, column=7, value="No POs mapped").alignment = left_align
            ws2.cell(row=row_num, column=8, value=primary_wk).alignment = center
            for c in range(1, 9):
                ws2.cell(row=row_num, column=c).border = thin_border
            ws2.row_dimensions[row_num].height = 30
            row_num += 1
            continue

        for po, data in sorted(mapped.items(), key=lambda x: _po_sort_key(x[0])):
            data = data or {}
            ws2.cell(row=row_num, column=1, value=co).alignment = center
            ws2.cell(row=row_num, column=2, value=po).alignment = center

            strength = data.get("strength")
            cell_s = ws2.cell(row=row_num, column=3, value=strength)
            cell_s.alignment = center
            if strength == 3:
                cell_s.fill = s3_fill
            elif strength == 2:
                cell_s.fill = s2_fill
            elif strength == 1:
                cell_s.fill = s1_fill

            x_pct = data.get("mapping_strength_pct")
            ws2.cell(
                row=row_num, column=4,
                value=x_pct if x_pct is not None else "–",
            ).alignment = center

            yes_c = data.get("yes_count")
            tot = data.get("total_pis")
            ratio = f"{yes_c}/{tot}" if yes_c is not None and tot is not None else "–"
            ws2.cell(row=row_num, column=5, value=ratio).alignment = center

            pis = ", ".join(str(p) for p in (data.get("contributing_pis") or []))
            ws2.cell(row=row_num, column=6, value=pis).alignment = center
            ws2.cell(row=row_num, column=7, value=data.get("justification", "")).alignment = left_align
            ws2.cell(row=row_num, column=8, value=primary_wk).alignment = center

            for c in range(1, 9):
                ws2.cell(row=row_num, column=c).border = thin_border
            ws2.row_dimensions[row_num].height = 50
            row_num += 1

    ws2.column_dimensions["A"].width = 8
    ws2.column_dimensions["B"].width = 8
    ws2.column_dimensions["C"].width = 10
    ws2.column_dimensions["D"].width = 10
    ws2.column_dimensions["E"].width = 14
    ws2.column_dimensions["F"].width = 28
    ws2.column_dimensions["G"].width = 70
    ws2.column_dimensions["H"].width = 12

    # ----- Sheet 4: Understanding & Recommendations -----
    ws3 = wb.create_sheet("4_Understanding_Recommendations")

    ws3["A1"] = "Deep Understanding of each CO + Recommendations"
    ws3["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws3.merge_cells("A1:D1")

    headers3 = ["CO", "Understanding of CO", "Suggested WKs", "Recommendations"]
    for col, h in enumerate(headers3, 1):
        ws3.cell(row=3, column=col, value=h)
    _style_header(ws3, 3, 1, 4)

    for i, res in enumerate(results):
        r = 4 + i
        ws3.cell(row=r, column=1, value=res.get("co_id", "")).alignment = center

        if "error" in res:
            ws3.cell(row=r, column=2, value=f"ERROR: {res.get('error', '')}").alignment = left_align
            ws3.cell(row=r, column=2).fill = error_fill
            ws3.cell(row=r, column=3, value="").alignment = center
            ws3.cell(row=r, column=4, value="").alignment = left_align
        else:
            ws3.cell(row=r, column=2, value=res.get("understanding", "")).alignment = left_align
            suggested = res.get("suggested_wks") or []
            ws3.cell(row=r, column=3, value=", ".join(str(w) for w in suggested)).alignment = center
            ws3.cell(row=r, column=4, value=res.get("recommendations", "")).alignment = left_align

        for c in range(1, 5):
            ws3.cell(row=r, column=c).border = thin_border
        ws3.row_dimensions[r].height = 70

    ws3.column_dimensions["A"].width = 8
    ws3.column_dimensions["B"].width = 65
    ws3.column_dimensions["C"].width = 25
    ws3.column_dimensions["D"].width = 50

    # ----- Sheet 5: PI Yes/No Detail (GAPC document style) -----
    # Mirrors tables on pages 51–56 of GAPC V4.0: for each PO, every PI marked Yes/No per CO
    ws_pi = wb.create_sheet("5_PI_YesNo_Detail")
    ws_pi["A1"] = "PO–Competency–PI–CO Mapping Status (GAPC V4.0 Step 6)"
    ws_pi["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws_pi.merge_cells("A1:G1")
    ws_pi["A2"] = (
        "For each PO, every Performance Indicator is marked Yes or No for each CO. "
        "X = (Yes count / Total PIs) × 100 → Mapping Value 1/2/3 per official rubrics."
    )
    ws_pi["A2"].font = Font(italic=True, size=9)
    ws_pi.merge_cells("A2:G2")

    co_ids = [res.get("co_id", f"CO{i+1}") for i, res in enumerate(results)]
    # Build lookup: co_id → po → set of yes PI ids
    yes_lookup: dict[str, dict[str, set]] = {}
    for res in results:
        cid = res.get("co_id", "")
        yes_lookup[cid] = {}
        if "error" in res:
            continue
        for po, data in (res.get("mapped_pos") or {}).items():
            yes_lookup[cid][po] = set(data.get("contributing_pis") or [])

    row = 4
    for po in ALL_POS:
        info = GAPC_PO_CATALOG.get(po)
        if not info:
            continue
        # Only emit tables for POs that at least one CO maps to (or always show all)
        any_mapped = any(po in yes_lookup.get(cid, {}) for cid in co_ids)
        if not any_mapped:
            continue

        ws_pi.cell(row=row, column=1, value=f"{po}: {info['title']}").font = Font(
            bold=True, size=11, color="1F4E79"
        )
        ws_pi.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3 + len(co_ids))
        row += 1
        ws_pi.cell(row=row, column=1, value=info.get("statement", "")).font = Font(
            italic=True, size=9
        )
        ws_pi.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3 + len(co_ids))
        row += 1

        # Header
        headers = ["Competency", "PI ID", "Performance Indicator"] + co_ids
        for col, h in enumerate(headers, 1):
            ws_pi.cell(row=row, column=col, value=h)
        _style_header(ws_pi, row, 1, len(headers))
        row += 1

        # PI rows
        for pid, pdesc in info["pis"].items():
            # Competency id is first two parts e.g. 1.1 from 1.1.1
            parts = pid.split(".")
            competency = ".".join(parts[:2]) if len(parts) >= 2 else pid
            ws_pi.cell(row=row, column=1, value=competency).alignment = center
            ws_pi.cell(row=row, column=2, value=pid).alignment = center
            ws_pi.cell(row=row, column=3, value=pdesc).alignment = left_align
            for j, cid in enumerate(co_ids):
                is_yes = pid in yes_lookup.get(cid, {}).get(po, set())
                cell = ws_pi.cell(row=row, column=4 + j, value="Yes" if is_yes else "No")
                cell.alignment = center
                if is_yes:
                    cell.fill = s3_fill
                else:
                    cell.fill = no_fill
            for c in range(1, 4 + len(co_ids)):
                ws_pi.cell(row=row, column=c).border = thin_border
            row += 1

        # Mapping Strength row
        ws_pi.cell(row=row, column=1, value="Mapping Strength X (%)").font = Font(bold=True)
        ws_pi.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        for j, cid in enumerate(co_ids):
            yes_set = yes_lookup.get(cid, {}).get(po, set())
            total = len(info["pis"])
            yes_c = len(yes_set)
            _, x_pct = gapc_mapping_value(yes_c, total)
            cell = ws_pi.cell(row=row, column=4 + j, value=x_pct)
            cell.alignment = center
            cell.font = Font(bold=True)
            cell.border = thin_border
        for c in range(1, 4):
            ws_pi.cell(row=row, column=c).border = thin_border
        row += 1

        # Mapping Value row
        ws_pi.cell(row=row, column=1, value="Mapping Value (1/2/3)").font = Font(bold=True)
        ws_pi.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        for j, cid in enumerate(co_ids):
            yes_set = yes_lookup.get(cid, {}).get(po, set())
            total = len(info["pis"])
            yes_c = len(yes_set)
            strength, _ = gapc_mapping_value(yes_c, total)
            cell = ws_pi.cell(row=row, column=4 + j, value=strength if strength else "–")
            cell.alignment = center
            cell.font = Font(bold=True)
            if strength == 3:
                cell.fill = s3_fill
            elif strength == 2:
                cell.fill = s2_fill
            elif strength == 1:
                cell.fill = s1_fill
            else:
                cell.fill = no_fill
            cell.border = thin_border
        for c in range(1, 4):
            ws_pi.cell(row=row, column=c).border = thin_border
        row += 2  # blank line between POs

    ws_pi.column_dimensions["A"].width = 14
    ws_pi.column_dimensions["B"].width = 10
    ws_pi.column_dimensions["C"].width = 70
    for j in range(len(co_ids)):
        ws_pi.column_dimensions[get_column_letter(4 + j)].width = 8

    # ----- Sheet 6: Summary -----
    ws4 = wb.create_sheet("6_Summary")
    ws4["A1"] = "Summary"
    ws4["A1"].font = Font(bold=True, size=13, color="1F4E79")

    covered = [
        po for po, st in stats["po_stats"].items() if st["total_mapped"] > 0
    ]
    ov = stats["overall"]

    ws4["A3"] = "POs covered:"
    ws4["B3"] = ", ".join(sorted(covered, key=_po_sort_key)) if covered else "None"
    ws4["A4"] = "Model used:"
    ws4["B4"] = model_name
    ws4["A5"] = "Total COs analyzed:"
    ws4["B5"] = ov["total_cos"]
    ws4["A6"] = "COs with errors:"
    ws4["B6"] = ov["errors"]
    ws4["A7"] = "Total CO-PO mappings:"
    ws4["B7"] = ov["total_mappings"]
    ws4["A8"] = "Overall average strength:"
    ws4["B8"] = ov["avg_strength"] if ov["avg_strength"] is not None else "–"
    ws4["A10"] = "Strength formula (GAPC V4.0):"
    ws4["B10"] = "X = (Yes PIs / Total PIs) × 100 → 10-33=1, 34-67=2, 68-100=3"
    ws4["A11"] = "Note:"
    ws4["A12"] = "Sheet 2 – Strength_Calculation: per-CO and per-PO averages and counts."
    ws4["A13"] = "Sheet 3 – Mapping_with_Justification: X%, Yes/Total, contributing PIs, justification."
    ws4["A14"] = "Sheet 5 – PI_YesNo_Detail: full Yes/No table per PI per CO (document Step 6 style)."
    if ov["errors"]:
        ws4["A16"] = "Warning:"
        ws4["B16"] = f"{ov['errors']} CO(s) failed analysis – see rows marked ERROR."

    ws4.column_dimensions["A"].width = 26
    ws4.column_dimensions["B"].width = 70

    wb.save(output_path)
    print(f"\nExcel report saved → {output_path}")


# ---------------------------------------------------------------------------
# Full CO-PO pipeline
# ---------------------------------------------------------------------------

def generate_co_po_mapping(
    model_name: str,
    api_key: str,
    input_excel: str = DEFAULT_INPUT_EXCEL,
    output_excel: str = DEFAULT_OUTPUT_EXCEL,
    sheet_name: Optional[str] = None,
) -> list[dict]:
    """
    Complete pipeline (GAPC V4.0):
      1. Initialize Gemini client
      2. Read COs from Excel
      3. Analyze each CO against Competencies & Performance Indicators
      4. Compute mapping strength X = (Yes PIs / Total PIs) × 100 → value 1/2/3
      5. Write multi-sheet Excel report + raw JSON

    Parameters
    ----------
    model_name : str
        **Required.** Gemini model id (e.g. "gemini-2.5-flash").
    api_key : str
        **Required.** Gemini API key.
    """
    if not model_name or not str(model_name).strip():
        raise ValueError("model_name is required (e.g. 'gemini-2.5-flash').")

    print("=" * 60)
    print("CO-PO Mapping Generator (GAPC V4.0 + Gemini)")
    print(f"Model: {model_name}")
    print("=" * 60)

    client = init_gemini_client(api_key)
    results = analyze_all_cos(
        input_excel, client, model_name=model_name, sheet_name=sheet_name
    )

    json_path = "gemini_raw_analysis.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"Raw analysis saved → {json_path}")

    create_excel_report(
        results,
        model_name=model_name,
        output_path=output_excel,
        input_excel=input_excel,
    )
    print("\nProcess completed successfully!")
    return results


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="CO-WK / CO-PO mapping for NBA Outcome-Based Education"
    )
    sub = parser.add_subparsers(dest="command")

    # PDF → WK pipeline
    p1 = sub.add_parser("wk", help="Generate CO + WK mapping from a syllabus PDF")
    p1.add_argument("pdf", help="Path to syllabus PDF")
    p1.add_argument("--model", required=True, help="Gemini model name (required), e.g. gemini-2.5-flash")
    p1.add_argument("--api-key", required=True, help="Gemini API key (required)")
    p1.add_argument("-o", "--output", default=DEFAULT_INPUT_EXCEL, help="Output Excel (default: CO_WK_Mapping.xlsx)")

    # Excel → PO pipeline
    p2 = sub.add_parser("po", help="Generate CO-PO mapping from an Excel of COs (GAPC V4.0)")
    p2.add_argument("excel", nargs="?", default=DEFAULT_INPUT_EXCEL)
    p2.add_argument("--model", required=True, help="Gemini model name (required), e.g. gemini-2.5-flash")
    p2.add_argument("--api-key", required=True, help="Gemini API key (required)")
    p2.add_argument("-o", "--output", default=DEFAULT_OUTPUT_EXCEL)
    p2.add_argument("--sheet", default=None)

    args = parser.parse_args()

    if args.command == "wk":
        generate_co_wk_excel(
            pdf_path=args.pdf,
            model_name=args.model,
            api_key=args.api_key,
            output_excel_path=args.output,
        )
    elif args.command == "po":
        generate_co_po_mapping(
            model_name=args.model,
            api_key=args.api_key,
            input_excel=args.excel,
            output_excel=args.output,
            sheet_name=args.sheet,
        )
    else:
        parser.print_help()

# Add this function at the very bottom of obe/mapper.py

def main_cli():
    """CLI Entry point for obe-mapper command."""
    import argparse

    parser = argparse.ArgumentParser(
        description="CO-WK / CO-PO mapping for NBA Outcome-Based Education"
    )
    sub = parser.add_subparsers(dest="command")

    p1 = sub.add_parser("wk", help="Generate CO + WK mapping from a syllabus PDF")
    p1.add_argument("pdf", help="Path to syllabus PDF")
    p1.add_argument("--model", required=True, help="Gemini model name")
    p1.add_argument("--api-key", required=True, help="Gemini API key")
    p1.add_argument("-o", "--output", default=DEFAULT_INPUT_EXCEL)

    p2 = sub.add_parser("po", help="Generate CO-PO mapping from Excel")
    p2.add_argument("excel", nargs="?", default=DEFAULT_INPUT_EXCEL)
    p2.add_argument("--model", required=True, help="Gemini model name")
    p2.add_argument("--api-key", required=True, help="Gemini API key")
    p2.add_argument("-o", "--output", default=DEFAULT_OUTPUT_EXCEL)
    p2.add_argument("--sheet", default=None)

    args = parser.parse_args()

    if args.command == "wk":
        generate_co_wk_excel(
            pdf_path=args.pdf,
            model_name=args.model,
            api_key=args.api_key,
            output_excel_path=args.output,
        )
    elif args.command == "po":
        generate_co_po_mapping(
            model_name=args.model,
            api_key=args.api_key,
            input_excel=args.excel,
            output_excel=args.output,
            sheet_name=args.sheet,
        )
    else:
        parser.print_help()
