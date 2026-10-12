import re
import os
import io
import pandas as pd
import streamlit as st
import pypdf
import docx
from pptx import Presentation
from google import genai
from google.genai import types

# ==============================================================================
# 1. PAGE CONFIGURATION & STYLING
# ==============================================================================
st.set_page_config(
    page_title="NIET Math Coaching Hub",
    page_icon="📐",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.5rem;
    }
    .glow-card {
        background-color: #f0fdf4;
        border-left: 5px solid #22c55e;
        padding: 1.2rem;
        border-radius: 8px;
        margin-bottom: 1rem;
    }
    .artifact-card {
        background-color: #ffffff;
        border: 2px solid #3b82f6;
        padding: 1.5rem;
        border-radius: 10px;
        margin-top: 1rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .role-badge {
        background-color: #eff6ff;
        border: 1px solid #bfdbfe;
        color: #1e40af;
        padding: 0.3rem 0.6rem;
        border-radius: 4px;
        font-weight: bold;
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. FERPA SANITIZATION & MULTI-FORMAT FILE PARSER ENGINE
# ==============================================================================
PII_COLUMN_KEYWORDS = [
    'first name', 'last name', 'student name', 'full name', 'student_name',
    'student id', 'local id', 'state id', 'student_id', 'id number', 'id',
    'dob', 'date of birth', 'birthdate', 'gender', 'ethnicity', 'race',
    'teacher', 'teacher name', 'instructor', 'email', 'phone'
]
PRESERVE_KEYWORDS = ['score', 'percent', 'pct', 'level', 'tier', 'question', 'item', 'raw', 'scale', 'q1', 'q2', 'q3', 'ans']

def sanitize_eduphoria_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    cols_to_drop = []
    for col in df.columns:
        col_lower = str(col).strip().lower()
        is_pii = any(pii_kw in col_lower for pii_kw in PII_COLUMN_KEYWORDS)
        is_metric = any(preserve_kw in col_lower for preserve_kw in PRESERVE_KEYWORDS)
        if is_pii and not is_metric:
            cols_to_drop.append(col)
            
    sanitized_df = df.drop(columns=cols_to_drop, errors='ignore').copy()
    if 'Student' in sanitized_df.columns:
        sanitized_df['Student'] = [f"Student {i+1}" for i in range(len(sanitized_df))]
    return sanitized_df, cols_to_drop

def sanitize_text_pii(text: str) -> str:
    if not text:
        return text
    text = re.sub(r'\b\d{6,10}\b', '[REDACTED_STUDENT_ID]', text)
    text = re.sub(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', '[REDACTED_EMAIL]', text)
    text = re.sub(r'\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b', '[REDACTED_PHONE]', text)
    return text

def extract_content_from_file(uploaded_file) -> tuple[str, types.Part | None]:
    file_name = uploaded_file.name.lower()
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)
    
    extracted_text = ""
    pdf_part = None
    
    if file_name.endswith('.pdf'):
        pdf_part = types.Part.from_bytes(data=file_bytes, mime_type="application/pdf")
        try:
            reader = pypdf.PdfReader(io.BytesIO(file_bytes))
            for i, page in enumerate(reader.pages):
                page_text = page.extract_text() or ""
                extracted_text += f"\n--- PAGE {i+1} ---\n{page_text}"
        except Exception:
            extracted_text = f"[PDF Content Loaded: {file_name}]"
            
    elif file_name.endswith(('.docx', '.doc')):
        try:
            doc = docx.Document(io.BytesIO(file_bytes))
            full_text = []
            for p in doc.paragraphs:
                if p.text.strip():
                    full_text.append(p.text)
            for table in doc.tables:
                for row in table.rows:
                    full_text.append(" | ".join([cell.text.strip() for cell in row.cells]))
            extracted_text = "\n".join(full_text)
        except Exception:
            extracted_text = f"[Word Document Content from {file_name}]"
            
    elif file_name.endswith(('.pptx', '.ppt')):
        try:
            prs = Presentation(io.BytesIO(file_bytes))
            slides_text = []
            for i, slide in enumerate(prs.slides):
                slide_content = [f"--- SLIDE {i+1} ---"]
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        slide_content.append(shape.text.strip())
                slides_text.append("\n".join(slide_content))
            extracted_text = "\n\n".join(slides_text)
        except Exception:
            extracted_text = f"[PowerPoint Presentation Content from {file_name}]"

    elif file_name.endswith(('.xlsx', '.xls')):
        try:
            excel_df = pd.read_excel(io.BytesIO(file_bytes))
            extracted_text = f"--- EXCEL SPREADSHEET DATA ---\n" + excel_df.to_markdown(index=False)
        except Exception:
            extracted_text = f"[Excel Data Content from {file_name}]"

    elif file_name.endswith('.csv'):
        try:
            csv_df = pd.read_csv(io.BytesIO(file_bytes))
            extracted_text = f"--- CSV DATA ---\n" + csv_df.to_markdown(index=False)
        except Exception:
            extracted_text = file_bytes.decode('utf-8', errors='ignore')

    else:
        extracted_text = file_bytes.decode('utf-8', errors='ignore')
        
    return extracted_text, pdf_part

# ==============================================================================
# 3. SIDEBAR CONFIGURATION & COMPLETE TEKS DATABASES
# ==============================================================================
TEKS_7TH_MATH = [
    "7.2A - Sets & Subsets of Rational Numbers",
    "7.3A - Operations with Rational Numbers",
    "7.3B - Apply Rational Number Operations (Real-World)",
    "7.4A - Constant Rate of Change",
    "7.4B - Calculate Unit Rates",
    "7.4C - Constant of Proportionality (k = y/x)",
    "7.4D - Solve Ratio, Rate & Percent Problems",
    "7.4E - Convert Measurement Systems",
    "7.5A - Critical Attributes of Similarity",
    "7.5B - Describe Pi (π) as Ratio of Circumference to Diameter",
    "7.5C - Similar Shapes & Scale Drawings",
    "7.6A - Represent Sample Spaces (Simple & Compound)",
    "7.6B - Select & Use Simulations for Probability",
    "7.6C - Experimental vs Theoretical Probability",
    "7.6D - Make Predictions using Experimental Data",
    "7.6E - Probabilities of Compound Events",
    "7.6F - Theoretical Probability of Events",
    "7.6G - Qualitative & Quantitative Probability Predictions",
    "7.6H - Solve Problems with Theoretical Probability",
    "7.6I - Determine Experimental & Theoretical Probabilities",
    "7.7A - Linear Relationships (Tables, Graphs, Equations y = mx + b)",
    "7.8A - Model Equations (a + b = c & ab = c)",
    "7.8B - Write One-Variable Two-Step Equations & Inequalities",
    "7.8C - Write Real-World Stories for Equations & Inequalities",
    "7.9A - Area of Composite Figures",
    "7.9B - Volume of Rectangular & Triangular Prisms & Pyramids",
    "7.9C - Circumference & Area of Circles",
    "7.10A - Write Two-Step Equations & Inequalities",
    "7.10B - Represent Inequality Solutions on Number Lines",
    "7.10C - Write Real-World Problems for Equations & Inequalities",
    "7.11A - Model & Solve Two-Step Equations & Inequalities",
    "7.11B - Verify Values that Make Equations/Inequalities True",
    "7.11C - Write & Solve Equations Using Geometric Concepts",
    "7.12A - Compare Data Sets (Dot Plots & Box Plots)",
    "7.12B - Make Inferences from Random Samples",
    "7.12C - Compare Populations Using Random Samples",
    "7.13A - Calculate Sales Tax & Income Tax",
    "7.13B - Calculate Simple & Compound Interest",
    "7.13C - Personal Budgets & Monthly Expenses",
    "7.13D - Analyze Net Worth Statements",
    "7.13E - Calculate Percentage Increase & Decrease",
    "7.13F - Monetary Incentives, Coupons & Rebates",
    "Custom TEKS"
]

TEKS_ALGEBRA_1 = [
    "A.2A - Domain & Range of Linear Functions",
    "A.2B - Write Linear Equations in Two Variables",
    "A.2C - Write Linear Equations (Table, Graph, Verbal)",
    "A.2D - Write & Solve Direct Variation Equations",
    "A.2E - Parallel & Perpendicular Line Equations",
    "A.2F - Line Equation Given Point & Slope",
    "A.2G - Line Equation Given Two Points",
    "A.2H - Write Linear Inequalities in Two Variables",
    "A.2I - Write Systems of Two Linear Equations",
    "A.3A - Determine Slope (Table, Graph, Two Points, Equation)",
    "A.3B - Calculate Rate of Change",
    "A.3C - Graph Linear Functions & Identify Key Features",
    "A.3D - Graph Linear Inequalities in Two Variables",
    "A.3E - Transformations on Parent Linear Function f(x) = x",
    "A.3F - Graph Systems of Two Linear Inequalities",
    "A.3G - Estimate Solutions to Systems of Linear Equations",
    "A.3H - Graph Systems of Two Linear Equations",
    "A.4A - Calculate Correlation Coefficient & Linear Regression",
    "A.4B - Direct Variation vs Linear Functions",
    "A.4C - Write Linear Functions to Model Real-World Data",
    "A.5A - Solve Linear Equations in One Variable",
    "A.5B - Solve Linear Inequalities in One Variable",
    "A.5C - Solve Systems of Two Linear Equations",
    "A.6A - Domain & Range of Quadratic Functions",
    "A.6B - Write Quadratic Functions (Vertex & Point)",
    "A.6C - Write Quadratic Functions in Vertex & Standard Form",
    "A.7A - Graph Quadratic Functions & Identify Key Features",
    "A.7B - Linear Factors & Zeros of Quadratic Functions",
    "A.7C - Transformations on Parent Quadratic Function f(x) = x²",
    "A.8A - Solve Quadratic Equations (Factoring, Square Roots
