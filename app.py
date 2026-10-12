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

st.markdown('''
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
''', unsafe_allow_html=True)

# ==============================================================================
# 2. SESSION STATE INITIALIZATION FOR CHAT
# ==============================================================================
if "chat_history" not in st.session_state:
    st.session_state.chat_history = [
        {
            "role": "assistant",
            "content": "👋 Hi! I am your NIET Math Coaching Assistant. Ask me to clarify any generated materials, adapt strategies for specific student populations, refine questions, or share feedback!"
        }
    ]

# ==============================================================================
# 3. FERPA SANITIZATION & MULTI-FORMAT FILE PARSER ENGINE
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
# 4. SIDEBAR CONFIGURATION & TEKS DATABASES
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
    "A.8A - Solve Quadratic Equations (Factoring/Square Roots/Formula)",
    "A.8B - Write & Solve Quadratic Equations in Real-World Contexts",
    "A.9A - Domain & Range of Exponential Functions",
    "A.9B - Interpret Parameters a & b in Exponential Functions",
    "A.9C - Write Exponential Functions (Growth & Decay)",
    "A.9D - Graph Exponential Functions & Identify Key Features",
    "A.9E - Write Exponential Functions from Tables/Graphs",
    "A.10A - Add & Subtract Polynomials",
    "A.10B - Multiply Polynomials",
    "A.10C - Quotient of Polynomials (Polynomial Division)",
    "A.10D - Rewrite Polynomial Expressions",
    "A.10E - Factor Trinomials (ax² + bx + c)",
    "A.10F - Factor Difference of Squares",
    "A.11A - Simplify Numerical Radical Expressions",
    "A.11B - Simplify Algebraic Expressions (Laws of Exponents)",
    "A.12A - Determine if a Relation is a Function",
    "A.12B - Evaluate Function Notation f(x)",
    "A.12C - Identify Terms of Arithmetic & Geometric Sequences",
    "A.12D - Formulas for Arithmetic & Geometric Sequences",
    "A.12E - Solve Formulas for a Specified Variable",
    "Custom TEKS"
]

NIET_ALL_INDICATORS = [
    "None",
    "Standards and Objectives",
    "Motivating Students",
    "Presenting Instructional Content",
    "Lesson Structure and Pacing",
    "Activities and Materials",
    "Questioning",
    "Academic Feedback",
    "Grouping Students",
    "Teacher Content Knowledge",
    "Teacher Knowledge of Students",
    "Thinking",
    "Problem Solving"
]

with st.sidebar:
    st.image("https://img.icons8.com/color/96/calculator.png", width=64)
    st.title("Coaching Settings")
    
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        api_key = st.text_input("Enter Gemini API Key (Optional)", type="password")
    
    st.divider()
    
    selected_grade = st.radio(
        "Select Target Course:",
        ["7th Grade Math", "9th Grade Algebra 1"]
    )
    
    teks_list = TEKS_7TH_MATH if selected_grade == "7th Grade Math" else TEKS_ALGEBRA_1
    selected_teks = st.selectbox("Select TEKS Focus:", teks_list)
    
    if selected_teks == "Custom TEKS":
        selected_teks = st.text_input("Enter Custom TEKS (e.g., 7.10B or A.12E):")

    st.divider()
    st.caption("🛡️ **FERPA Privacy Active:** In-memory PII sanitization enabled.")

client = genai.Client(api_key=api_key) if api_key else None

# ==============================================================================
# 5. HEADER
# ==============================================================================
st.markdown(f'''
<div class="main-header">
    <h1 style="margin:0;">📐 NIET Math Coaching Intelligence Hub</h1>
    <p style="margin:0; opacity:0.9;">Tailored for <b>{selected_grade}</b> • TEKS-Adaptive CRA + Complex Instruction + Accountable Talk Moves • STAAR 2.0 Dual Alignment</p>
</div>
''', unsafe_allow_html=True)

# ==============================================================================
# 6. MAIN TABS
# ==============================================================================
tab_elevator, tab_assess, tab_data, tab_cluster, tab_chat = st.tabs([
    "✨ NIET Lesson Elevator",
    "📝 Warmup & Exit Ticket Generator",
    "📊 Eduphoria Data & Multimodal Misconceptions",
    "👥 PLC Cluster & Reflective Inquiry Coach",
    "💬 Interactive Chat & Feedback"
])

# ------------------------------------------------------------------------------
# TAB 1: NIET LESSON ELEVATOR (WITH ENHANCED FEEDBACK DIRECTIVES)
# ------------------------------------------------------------------------------
with tab_elevator:
    st.subheader("Elevate Lesson Draft or File Upload to NIET Level 5")
    st.caption("Upload a PDF, Word Doc, Google Slides export (.pptx), Excel/Sheets (.xlsx), or TXT file to receive detailed NIET rubric feedback and an elevated lesson plan.")

    col_e1, col_e2 = st.columns([1, 1])
    
    with col_e1:
        primary_indicator = st.selectbox(
            "Primary NIET Indicator Focus:",
            [ind for ind in NIET_ALL_INDICATORS if ind != "None"],
            index=5
        )
        
        secondary_indicator = st.selectbox(
            "Secondary NIET Indicator Focus (Optional):",
            NIET_ALL_INDICATORS,
            index=0
        )
        
        uploaded_lesson_file = st.file_uploader(
            "Upload Lesson File (PDF, Word, Slides, Sheets, or TXT):",
            type=["pdf", "docx", "doc", "pptx", "ppt", "xlsx", "xls", "csv", "txt"],
            key="lesson_file_upload"
        )
        
        lesson_text_input = st.text_area(
            "Or paste lesson text / objective here:",
            height=140,
            placeholder=f"Example for {selected_grade}: Objective is to solve linear equations or two-step inequalities."
        )

        btn_elevate = st.button("🚀 Generate Level 5 NIET Coaching Feedback Brief", type="primary", use_container_width=True)

    with col_e2:
        if btn_elevate:
            if not lesson_text_input and not uploaded_lesson_file:
                st.warning("⚠️ Please paste lesson text or upload a document to generate feedback.")
            else:
                clean_text = sanitize_text_pii(lesson_text_input)
                file_text_content = ""
                pdf_part = None
                
                if uploaded_lesson_file:
                    file_text_content, pdf_part = extract_content_from_file(uploaded_lesson_file)
                    st.info(f"📄 Loaded '{uploaded_lesson_file.name}' successfully.")
                
                combined_lesson_content = f"{clean_text}\n\n{file_text_content}".strip()
                
                with st.spinner("Analyzing lesson draft against NIET Level 5 Indicators & generating coaching brief..."):
                    if client:
                        prompt_contents = []
                        if pdf_part:
                            prompt_contents.append(pdf_part)
                        
                        sec_clause = f"and Secondary Focus: '{secondary_indicator}'" if secondary_indicator != "None" else ""
                        
                        p1 = f"You are a Master Junior High Math Instructional Coach for {selected_grade} specializing in the NIET (National Instruction for Effective Teaching) rubric."
                        p2 = f"PRIMARY NIET INDICATOR FOCUS: '{primary_indicator}' {sec_clause}."
                        p3 = f"TARGET TEKS: {selected_teks} ({selected_grade})."
                        
                        p_rules = (
                            "YOU MUST PROVIDE A FULL 4-SECTION DETAILED COACHING BRIEF:\n\n"
                            "### 🌟 Section 1: Authentic Affirming Glows\n"
                            "- Detail 2-3 specific instructional strengths in the submitted material aligned to the target TEKS and NIET indicator.\n\n"
                            "### 🎯 Section 2: NIET Level 5 Coaching Grows & Strategic Refinements\n"
                            "- Provide actionable, direct coaching feedback on how to elevate this lesson from Level 3/4 to Level 5.\n"
                            "- Address student inquiry, cognitive demand, and proactive misconception prevention.\n\n"
                            "### 🗣️ Section 3: Verbatim Complex Instruction & Accountable Talk Script\n"
                            "- Provide 2-3 specific Teacher Coaching Cues and Verbatim Classroom Scripts incorporating Accountable Talk moves (Revoicing, Reason-Checking, Adding On).\n\n"
                            "### 📜 Section 4: Elevated POGIL Discovery Activity Mat\n"
                            "- Present the visual Markdown discovery mat for student teams.\n"
                            "- Include a standalone ```text ``` block with the complete elevated lesson material for easy copying.\n\n"
                            "FORMATTING RULES:\n"
                            "1. NO HTML TAGS: Do NOT use raw HTML tags (e.g., <br>, <div>).\n"
                            "2. NO MARKDOWN TABLES FOR COMPLEX MATH: Use bold section headers (####), bullet points, and blockquotes.\n"
                            "3. CURRENCY: Format money as 'USD $10' in prose, and equations in LaTeX ($3g + 10 < 40$).\n"
                        )
                        system_prompt = f"{p1}\n{p2}\n{p3}\n\n{p_rules}"
                        prompt_contents.append(f"{system_prompt}\n\nSubmitted Lesson Content:\n{combined_lesson_content}")
                        
                        response = client.models.generate_content(
                            model="gemini-1.5-pro",
                            contents=prompt_contents
                        )
                        st.markdown(response.text)
                    else:
                        # Rich Offline Fallback Coaching Brief
                        st.markdown(f"""
                        ### 🌟 Section 1: Authentic Affirming Glows
                        * **Clear Focus:** The lesson draft establishes a solid foundational objective aligned to **{selected_teks}**.
                        * **Real-World Hook:** The context provides students with a recognizable starting scenario.
                        
                        ### 🎯 Section 2: NIET Level 5 Coaching Grows ({primary_indicator})
                        * **Elevate Inquiry & Student Ownership:** Shift teacher-led direct instruction into a student-led POGIL model where students discover the math rule before formal notation.
                        * **Embed Cognitive Scaffolding:** Add visual CRA (Concrete-Representational-Abstract) models to prevent the common misconception of misinterpreting fixed vs variable rates.
                        * **Differentiate Questioning:** Transition recall questions into open-ended justification prompts (e.g., *"How does changing the fixed fee alter our boundary limit?"*).

                        ### 🗣️ Section 3: Verbatim Accountable Talk Scripting
                        * **Teacher Coaching Cue:** When students make an initial estimate, use *Revoicing* to enforce precise mathematical terminology.
                        * **Verbatim Script:** *"I hear Alex saying that the $10 setup fee is added only once. Who can add on to why that fee isn't multiplied by the number of months?"*

                        ### 📜 Section 4: Elevated POGIL Discovery Activity
                        #### 🟦 STUDENT INQUIRY MAT: {selected_teks}
                        1. **Exploratory Model:** Examine the scenario and compare a flat fee vs. a monthly rate.
                        2. **Group Critique:** Test values $x = 1, 2, 3$ to verify which side of the inequality remains true.
                        """)

# ------------------------------------------------------------------------------
# TAB 2: WARMUP & EXIT TICKET GENERATOR
# ------------------------------------------------------------------------------
with tab_assess:
    st.subheader("Targeted Assessment & Lesson Objective Generator")
    st.caption("Upload a lesson file/activity to automatically extract its core objective and target misconception. Generates a Warmup Launch into the lesson activity and an Exit Ticket assessing objective mastery.")

    col_a1, col_a2 = st.columns([1, 1])
    with col_a1:
        uploaded_assess_file = st.file_uploader(
            "Upload Lesson File / Plan (PDF, Word, Slides, Sheets, TXT):",
            type=["pdf", "docx", "doc", "pptx", "ppt", "xlsx", "xls", "csv", "txt"],
            key="assess_file_upload"
        )
        
        misconception_input = st.text_input(
            "Or manually enter / override Lesson Objective or Misconception:",
            value="",
            placeholder="e.g., Objective: Write two-step inequalities from real-world mobile data plans."
        )
        
        btn_gen_assess = st.button("⚡ Extract Objective & Build Launch & Exit Ticket", type="primary", use_container_width=True)

    with col_a2:
        if btn_gen_assess:
            if not uploaded_assess_file and not misconception_input.strip():
                st.warning("⚠️ Please upload a lesson file or enter an objective/misconception to continue.")
            else:
                clean_text = sanitize_text_pii(misconception_input)
                file_text_content = ""
                pdf_part = None
                
                if uploaded_assess_file:
                    file_text_content, pdf_part = extract_content_from_file(uploaded_assess_file)
                    st.info(f"📄 Loaded '{uploaded_assess_file.name}' for objective and misconception extraction.")
                
                combined_assess_input = f"Manual Input: {clean_text}\n\nUploaded Lesson Content:\n{file_text_content}".strip()
                
                with st.spinner("Extracting lesson objective and generating Lesson Launch Warmup and Objective Exit Ticket..."):
                    if client:
                        prompt_contents = []
                        if pdf_part:
                            prompt_contents.append(pdf_part)
                        
                        p1 = f"You are a STAAR 2.0 Assessment Specialist and Math Instructional Coach for {selected_grade}."
                        p2 = f"TARGET TEKS: {selected_teks} ({selected_grade})."
                        p_rules = (
                            "TASK:\n"
                            "1. Analyze the uploaded lesson content/file or text.\n"
                            "2. Extract and explicitly state:\n"
                            "   - Extracted Lesson Objective: Core target taught in the uploaded lesson/activity.\n"
                            "   - Target Misconception / Prerequisite Gap: Primary cognitive hurdle or prerequisite gap.\n"
                            "3. Generate a 5-minute Entrance Warmup as a LESSON LAUNCH (activates prior knowledge / hooks students into the upcoming activity).\n"
                            "4. Generate a 5-minute Exit Ticket as an OBJECTIVE MASTERY ASSESSMENT (directly evaluates mastery of the exact objective in the uploaded activity).\n\n"
                            "MANDATORY 4-OPTION MULTIPLE CHOICE RULE FOR BOTH WARMUP & EXIT TICKET:\n"
                            "Every item MUST feature exactly 4 Multiple-Choice options (Option A, Option B, Option C, Option D):\n"
                            "- 1 Correct Answer\n"
                            "- 2 Procedural / Conceptual Distractors (common STAAR errors)\n"
                            "- 1 Instant-Elimination Distractor (removable with ZERO computation using vocabulary or estimation - magnitude too big/small)\n\n"
                            "FORMATTING DIRECTIVES:\n"
                            "1. NO HTML TAGS: Banned raw HTML (no <br>, <b>, <div>). Use clean Markdown section headers (####) and standard line breaks.\n"
                            "2. NO MARKDOWN TABLES FOR QUESTION ITEMS: Present student items using clear bold section boxes and bullet points.\n"
                            "3. CURRENCY VS MATH: Write currency explicitly as 'USD $10' in prose, but math equations in LaTeX like '$3g + 10 < 40$'.\n\n"
                            "OUTPUT STRUCTURE REQUIRED:\n"
                            "### 🎯 Extracted Lesson Objective & TEKS Focus\n"
                            "* **Objective:** [1-sentence clear objective derived from the uploaded activity]\n"
                            "* **Primary Misconception / Gap:** [1-2 sentence sharp explanation]\n\n"
                            "### 📝 Entrance Warmup: Lesson Launch\n"
                            "*Purpose: Activates prerequisite concepts and hooks students into today's activity.*\n"
                            "#### 🟦 STUDENT-FACING LAUNCH QUESTION\n"
                            "**Context:** [Real-world teen scenario connecting to the upcoming activity]\n"
                            "**Task:** [Question prompt]\n"
                            "* [ ] **Option A:** [Text/math]\n"
                            "* [ ] **Option B:** [Text/math]\n"
                            "* [ ] **Option C:** [Text/math]\n"
                            "* [ ] **Option D:** [Text/math]\n\n"
                            "#### 🟩 METACOGNITIVE DEFENSE & DISTRACTOR CRITIQUE\n"
                            "1. **Instant Elimination Task:** Identify option removable immediately via vocabulary or estimation.\n"
                            "* **Sentence Stem:** *\"Option [X] can be eliminated immediately because...\"*\n"
                            "2. **Visual CRA Launch Task:** [Task setting up visual model for the activity]\n\n"
                            "### 🚪 Exit Ticket: Objective Mastery Assessment\n"
                            "*Purpose: Evaluates student mastery of the exact objective in the uploaded activity.*\n"
                            "#### 🟦 STUDENT-FACING ASSESSMENT QUESTION\n"
                            "**Context:** [Real-world teen scenario evaluating full objective application from the activity]\n"
                            "**Task:** [Question prompt]\n"
                            "* [ ] **Option A:** [Text/math]\n"
                            "* [ ] **Option B:** [Text/math]\n"
                            "* [ ] **Option C:** [Text/math]\n"
                            "* [ ] **Option D:** [Text/math]\n\n"
                            "#### 🟩 METACOGNITIVE DEFENSE & DISTRACTOR CRITIQUE\n"
                            "1. **Instant Elimination Task:** Identify option unreasonable based on vocabulary/estimation.\n"
                            "* **Sentence Stem:** *\"Option [X] is unreasonable because...\"*\n"
                            "2. **Objective Defense:** Explain why chosen option proves mastery of today's objective.\n\n"
                            "### ⚠️ Eduphoria Distractor Analysis Key\n"
                            "* **Warmup Launch Key:** Option A (Correct), Distractor 1 (Procedural), Distractor 2 (Procedural), Distractor 3 (Instant Elimination via vocab/estimation).\n"
                            "* **Exit Ticket Key:** Option A (Correct), Distractor 1 (Procedural), Distractor 2 (Procedural), Distractor 3 (Instant Elimination via vocab/estimation).\n"
                        )
                        system_prompt = f"{p1}\n{p2}\n\n{p_rules}"
                        prompt_contents.append(f"{system_prompt}\n\nInput Content:\n{combined_assess_input}")
                        
                        response = client.models.generate_content(
                            model="gemini-1.5-pro",
                            contents=prompt_contents
                        )
                        st.markdown(response.text)
                    else:
                        st.subheader(f"Generated Assessment Slips: {selected_teks}")
                        st.markdown("""
                        ### 🎯 Extracted Lesson Objective & TEKS Focus
                        * **Objective:** Students will model and solve two-step inequalities representing real-world gaming subscription budgets and verify boundary limits.
                        * **Primary Misconception / Gap:** Students confuse one-time setup fees with recurring per-item rates when constructing multi-step inequalities.
                        
                        ---

                        ### 📝 Entrance Warmup: Lesson Launch
                        *Purpose: Activates prerequisite concepts and hooks students into today's activity.*

                        #### 🟦 STUDENT-FACING LAUNCH QUESTION
                        **Context:** Jordan is setting up a mobile cloud pass. The pass requires a **USD $10** one-time activation fee plus **USD $3** per game downloaded.
                        
                        **Task:** Which visual tape diagram structure correctly launches our model for spending under USD $40?
                        * [ ] **Option A:** $3g + 10 < 40$
                        * [ ] **Option B:** $10g + 3 < 40$
                        * [ ] **Option C:** $3g + 10 > 40$
                        * [ ] **Option D:** $3g + 10 < 4$

                        #### 🟩 METACOGNITIVE DEFENSE & DISTRACTOR CRITIQUE
                        1. **Instant Elimination Task:** Identify which option can be eliminated immediately without heavy calculation based on estimation.
                        * **Sentence Stem:** *"Option D can be eliminated immediately because a spending limit of $4 is less than the $10 setup fee alone..."*
                        2. **Visual CRA Launch Task:** Sketch a visual balance scale comparing the fixed $10 setup fee to the variable $3 downloads.

                        ---

                        ### 🚪 Exit Ticket: Objective Mastery Assessment
                        *Purpose: Evaluates student mastery of the exact objective taught in the uploaded activity.*

                        #### 🟦 STUDENT-FACING ASSESSMENT QUESTION
                        **Context:** Jordan solved $-3x + 4 > 16$ to determine his remaining cloud storage allowance in gigabytes.
                        
                        **Task:** Select the correct boundary solution representing his storage allowance:
                        * [ ] **Option A:** $x < -4$
                        * [ ] **Option B:** $x > -4$
                        * [ ] **Option C:** $x < 4$
                        * [ ] **Option D:** $x < 400$

                        #### 🟩 METACOGNITIVE DEFENSE & DISTRACTOR CRITIQUE
                        1. **Instant Elimination Task:** Identify which option is unreasonable based on magnitude estimation.
                        * **Sentence Stem:** *"Option D can be eliminated immediately because 400 gigabytes is far beyond our monthly account scale..."*
                        2. **Objective Defense:** Explain how testing $x=0$ verifies that dividing by $-3$ reverses the inequality symbol.
                        
                        ---

                        ### ⚠️ Eduphoria Distractor Analysis Key
                        * **Warmup Launch Key:** Option A (Correct: $3g + 10 < 40$), Option B (Swapped rate with flat fee), Option C (Flipped inequality symbol), Option D (Instant Elimination: budget limit of $4 is impossible).
                        * **Exit Ticket Key:** Option A (Correct: $x < -4$), Option B (Forgot sign flip), Option C (Sign arithmetic error), Option D (Instant Elimination: magnitude of 400 is unreasonable).
                        """)

# ------------------------------------------------------------------------------
# TAB 3: EDUPHORIA DATA & MULTIMODAL MISCONCEPTION COACH
# ------------------------------------------------------------------------------
with tab_data:
    st.subheader("Eduphoria Student Responses & Test File Analysis")
    st.caption("Upload your Eduphoria student item response CSV along with the Test file.")

    st.info("🛡️ **FERPA Privacy Shield Active:** In-memory PII anonymization active.")

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        uploaded_csv = st.file_uploader("1. Upload Eduphoria Student Responses (.csv):", type=["csv"], key="eduphoria_csv")
    with col_u2:
        uploaded_pdf = st.file_uploader("2. Upload Assessment File (PDF, Word, Slides):", type=["pdf", "docx", "pptx", "txt"], key="test_pdf")

    btn_analyze_test = st.button("🔍 Analyze Test & Identify Top 4 Misconceptions", type="primary", use_container_width=True)

    if btn_analyze_test:
        if not uploaded_csv or not uploaded_pdf:
            st.warning("⚠️ Please upload BOTH the CSV and Assessment file to run the analysis.")
        else:
            try:
                raw_df = pd.read_csv(uploaded_csv)
                sanitized_df, dropped_pii = sanitize_eduphoria_dataframe(raw_df)
                csv_summary_str = sanitized_df.to_csv(index=False)
                
                test_file_text, pdf_part = extract_content_from_file(uploaded_pdf)
                
                st.success(f"🔒 FERPA Scrub Complete! Loaded '{uploaded_pdf.name}'. Analyzing items...")

                if client:
                    with st.spinner("Cross-referencing test questions with distractors..."):
                        p1 = f"You are a Master Junior High Math Instructional Coach for {selected_grade}."
                        p2 = f"TARGET TEKS: {selected_teks} ({selected_grade})."
                        p_rules = (
                            "TASK: Identify the TOP 4 MOST CRITICAL COMMON MISCONCEPTIONS across student responses.\n"
                            "For each misconception provide Target Item & TEKS, 1-2 sentence explanation, and Exploratory Addressing Strategy."
                        )
                        system_prompt = f"{p1}\n{p2}\n\n{p_rules}"
                        prompt_contents = []
                        if pdf_part:
                            prompt_contents.append(pdf_part)
                        
                        prompt_contents.append(f"{system_prompt}\n\nTest Content:\n{test_file_text}\n\nStudent Response Data:\n{csv_summary_str}")
                        
                        response = client.models.generate_content(
                            model="gemini-1.5-pro",
                            contents=prompt_contents
                        )
                        st.markdown(response.text)
                else:
                    st.markdown("### 📊 Sample Preview")
            except Exception as e:
                st.error(f"Error processing files: {e}")

# ------------------------------------------------------------------------------
# TAB 4: PLC CLUSTER & REFLECTIVE INQUIRY COACH
# ------------------------------------------------------------------------------
with tab_cluster:
    st.subheader("1-on-1 Coaching Debrief & Reflective Inquiry Coach")
    st.caption("Structures post-observation conferences using the Reflective Inquiry Protocol.")

    col_c1, col_c2 = st.columns([1, 1])
    with col_c1:
        observation_notes = st.text_area(
            "Paste Observation Notes or Growth Goal:",
            height=160,
            value=f"Teacher led a guided lesson on {selected_teks}."
        )
        btn_cluster = st.button("📋 Generate Reflective Inquiry Debrief Guide", type="primary", use_container_width=True)

    with col_c2:
        if btn_cluster:
            clean_notes = sanitize_text_pii(observation_notes)
            if client:
                system_prompt = (
                    "You are a Master Instructional Coach guiding a 1-on-1 post-observation teacher conference.\n"
                    "DEBRIEF PROTOCOL: REFLECTIVE INQUIRY (NON-DIRECTIVE DISCOVERY)\n"
                    "1. AFFIRMATION\n2. GUIDED REFLECTION (SEE IT)\n3. CO-CONSTRUCTED ACTION STEP (NAME IT)\n4. PRACTICE & COMMITMENT (DO IT)"
                )
                response = client.models.generate_content(
                    model="gemini-1.5-pro",
                    contents=f"{system_prompt}\n\nObservation Notes:\n{clean_notes}"
                )
                st.markdown(response.text)
            else:
                st.markdown("<div class='glow-card'><h4>🌟 Affirming Opening</h4></div>", unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# TAB 5: INTERACTIVE CHAT & FEEDBACK ASSISTANT
# ------------------------------------------------------------------------------
with tab_chat:
    st.subheader("💬 Interactive Coaching Chat & Feedback")
    st.caption("Ask follow-up questions, request modifications to generated outputs, adapt items for specific accommodations, or submit feedback.")

    # Header controls
    col_chat_hdr1, col_chat_hdr2 = st.columns([5, 1])
    with col_chat_hdr2:
        if st.button("🗑️ Clear Chat", type="secondary", use_container_width=True):
            st.session_state.chat_history = [
                {
                    "role": "assistant",
                    "content": "👋 Chat cleared! How can I assist you with your math coaching plans today?"
                }
            ]
            st.rerun()

    # Display chat history
    for message in st.session_state.chat_history:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    # Interactive Chat Input
    if chat_prompt := st.chat_input("Ask a question, request adjustments to a lesson/exit ticket, or leave feedback..."):
        clean_chat_prompt = sanitize_text_pii(chat_prompt)
        
        # Store & Display User Message
        st.session_state.chat_history.append({"role": "user", "content": clean_chat_prompt})
        with st.chat_message("user"):
            st.markdown(clean_chat_prompt)

        # Generate Assistant Response
        with st.chat_message("assistant"):
            if client:
                with st.spinner("Analyzing request..."):
                    # Build System Context with active TEKS & Grade
                    system_context = (
                        f"You are a Master Junior High Math Instructional Coach for {selected_grade}.\n"
                        f"ACTIVE COURSE: {selected_grade}\n"
                        f"ACTIVE TEKS FOCUS: {selected_teks}\n\n"
                        "DIRECTIVES:\n"
                        "1. Assist teachers and coaches with clarifying generated materials, scaffolding, accommodations, or adjusting questions.\n"
                        "2. Respond with actionable, peer-to-peer instructional wisdom.\n"
                        "3. Use LaTeX formatting for formal math expressions (e.g., $2x + 5 = 15$).\n"
                        "4. Do NOT use raw HTML tags."
                    )
                    
                    # Format recent history for model prompt
                    history_context_str = ""
                    for msg in st.session_state.chat_history[-6:]:
                        role_title = "User" if msg["role"] == "user" else "Coach Assistant"
                        history_context_str += f"\n\n{role_title}: {msg['content']}"

                    full_prompt = f"{system_context}\n\nConversation History:{history_context_str}"
                    
                    response = client.models.generate_content(
                        model="gemini-1.5-pro",
                        contents=full_prompt
                    )
                    
                    ai_response_text = response.text
                    st.markdown(ai_response_text)
                    st.session_state.chat_history.append({"role": "assistant", "content": ai_response_text})
            else:
                fallback_response = (
                    "Thank you for your inquiry/feedback! When connected to a valid Gemini API key, "
                    "I will generate adaptive answers, lesson tweaks, and real-time coaching support."
                )
                st.markdown(fallback_response)
                st.session_state.chat_history.append({"role": "assistant", "content": fallback_response})
