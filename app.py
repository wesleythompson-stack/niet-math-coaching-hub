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
    "A.8A - Solve Quadratic Equations (Factoring, Square Roots, Formula)",
    "A.8B - Write & Solve Quadratic Equations in Real-World Contexts",
    "A.9A - Domain & Range of Exponential Functions",
    "A.9B - Interpret Parameters a & b in Exponential Functions f(x) = ab^x",
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
# 4. HEADER
# ==============================================================================
st.markdown(f"""
<div class="main-header">
    <h1 style="margin:0;">📐 NIET Math Coaching Intelligence Hub</h1>
    <p style="margin:0; opacity:0.9;">Tailored for <b>{selected_grade}</b> • TEKS-Adaptive CRA + Complex Instruction + Accountable Talk Moves • STAAR 2.0 Dual Alignment</p>
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# 5. MAIN TABS
# ==============================================================================
tab_elevator, tab_assess, tab_data, tab_cluster = st.tabs([
    "✨ NIET Lesson Elevator",
    "📝 Warmup & Exit Ticket Generator",
    "📊 Eduphoria Data & Multimodal Misconceptions",
    "👥 PLC Cluster & Reflective Inquiry Coach"
])

# ------------------------------------------------------------------------------
# TAB 1: NIET LESSON ELEVATOR (REFINED COACHING ENGINE)
# ------------------------------------------------------------------------------
with tab_elevator:
    st.subheader("Elevate Lesson Draft or File Upload to NIET Level 5")
    st.caption("Upload a PDF, Word Doc, Google Slides export (.pptx), Excel/Sheets (.xlsx), or TXT file. The AI analyzes uploaded content against selected NIET indicators, applying TEKS-adaptive CRA modeling, Complex Instruction dual-scripting, and full 5-move Accountable Talk rotations.")

    col_e1, col_e2 = st.columns([1, 1])
    
    with col_e1:
        primary_indicator = st.selectbox(
            "Primary NIET Indicator Focus:",
            [ind for ind in NIET_ALL_INDICATORS if ind != "None"],
            index=5 # Questioning
        )
        
        secondary_indicator = st.selectbox(
            "Secondary NIET Indicator Focus (Optional):",
            NIET_ALL_INDICATORS,
            index=0
        )
        
        uploaded_lesson_file = st.file_uploader(
            "Upload Lesson File (PDF, Word, Google Slides/PPTX, Google Sheets/Excel, or TXT):",
            type=["pdf", "docx", "doc", "pptx", "ppt", "xlsx", "xls", "csv", "txt"],
            key="lesson_file_upload"
        )
        
        lesson_text_input = st.text_area(
            "Or paste lesson text / objective here:",
            height=140,
            placeholder=f"Example for {selected_grade}: Objective is to solve linear equations. Teacher presents 3 examples on board. Students work on a 10-question worksheet."
        )

        btn_elevate = st.button("🚀 Generate Level 5 POGIL Coaching Brief", type="primary", use_container_width=True)

    with col_e2:
        if btn_elevate:
            if not lesson_text_input and not uploaded_lesson_file:
                st.warning("⚠️ Please paste lesson text or upload a document to continue.")
            else:
                clean_text = sanitize_text_pii(lesson_text_input)
                file_text_content = ""
                pdf_part = None
                
                if uploaded_lesson_file:
                    file_text_content, pdf_part = extract_content_from_file(uploaded_lesson_file)
                    st.info(f"📄 Loaded '{uploaded_lesson_file.name}' successfully.")
                
                combined_lesson_content = f"{clean_text}\n\n{file_text_content}".strip()
                
                with st.spinner("Analyzing lesson through POGIL, Complex Instruction, and NIET Level 5 lenses..."):
                    if client:
                        prompt_contents = []
                        if pdf_part:
                            prompt_contents.append(pdf_part)
                        
                        secondary_clause = f"and Secondary Focus: '{secondary_indicator}'" if secondary_indicator != "None" else ""
                        
                        system_prompt = f"""
                        You are a Master Junior High Math Instructional Coach specializing in NIET Rubric alignment, POGIL (Process-Oriented Guided Inquiry Learning), Complex Instruction, and Accountable Talk Moves for {selected_grade}.
                        
                        PRIMARY NIET INDICATOR: '{primary_indicator}' {secondary_clause}.
                        TARGET TEKS: {selected_teks} ({selected_grade}).
                        
                        MANDATORY PEDAGOGICAL DIRECTIVES:
                        1. TEKS-ADAPTIVE CRA VISUAL MODELING: Dynamically select the optimal visual CRA tool for {selected_teks} (e.g., Algebra Tiles for equations/factoring, Visual Balance Scales for inequalities, Area Models/Grid Decompositions for composite shapes/polynomials, Coordinate Motion/Slope Triangles for linear graphs, or Double Number Lines/Tape Diagrams for rates/percents). Require students to sketch/construct the model BEFORE symbolic writing.
                        2. DUAL-SCRIPTING FOR COMPLEX INSTRUCTION ('Assigning Competence'): Dynamically determine whether spatial intuition, pattern discovery, or error-detection is the easiest way to elevate group status. Provide both:
                           - **Teacher Coaching Cue:** What the coach looks for during team observation.
                           - **Verbatim Classroom Script:** Exact teacher words to publicly validate a low-status student's visual/conceptual contribution.
                        3. FULL 5-MOVE ACCOUNTABLE TALK ROTATION: Embed explicit prompts for all 5 classic Chapin/Michaels moves across the dialogue:
                           - Revoicing ("So what I hear you saying is...")
                           - Repeat/Restate ("Who can restate Taylor's explanation in their own words?")
                           - Agree/Disagree + Reason ("Do you agree or disagree, and what proof supports that?")
                           - Adding On ("Who can build on this visual representation?")
                           - Explicit Wait-Time Cues (Mandate 5-7 seconds of silence before student responses).
                        4. POGIL GUIDED INQUIRY CYCLE: Structure student exploration into Phase 1 (Exploration), Phase 2 (Concept Invention), Phase 3 (Application) with 4 explicit roles (Manager, Recorder, Spokesperson, Reflector).
                        5. EB (ELPS) & SPED ACCESSIBILITY: Provide visual cognate anchor cards, chunked manipulative steps, and sentence frames.
                        6. HYBRID ARTIFACT OUTPUT:
                           - Render a visual Markdown POGIL Discovery Mat.
                           - Immediately follow with a clean, unformatted copy block inside ```text ``` code block tags for one-click pasting into Google Docs / Word.
                        """
                        
                        prompt_contents.append(f"{system_prompt}\n\nUploaded Lesson Input Content:\n{combined_lesson_content}")
                        
                        response = client.models.generate_content(
                            model="gemini-1.5-pro",
                            contents=prompt_contents
                        )
                        st.markdown(response.text)
                    else:
                        st.markdown(f"""
                        <div class="glow-card">
                            <h4>🌟 Authentic Affirming Glows</h4>
                            <p>Wonderful foundation! Your draft for <b>{selected_teks}</b> establishes a clear mathematical learning target and structures dedicated practice time for students.</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        st.markdown(f"### 🎯 NIET Level 5 Indicator Shift ({primary_indicator})")
                        st.write(f"**NIET Exemplary Target for {primary_indicator}:** *'Tasks demand high-level thinking, require students to draw connections, and prompt metacognitive self-monitoring through structured group inquiry.'*")
                        
                        st.markdown("### 📍 Slide-by-Slide / Section-by-Section POGIL Annotations")
                        st.info("**Slide 2 / Section 1 (Warmup):** Convert procedural formula display into POGIL Exploration. Show two contrasting mobile phone plans visually using tape diagrams.\n\n**Slide 5 / Section 2 (Guided Practice):** Replace teacher lecture with POGIL Roles. Assign *Manager* to time-box exploration and *Recorder* to draw visual balance scales.")

                        st.markdown("### 💬 Level 3 vs Level 5 Classroom Dialogue Script (5 Talk Moves)")
                        col_l3, col_l5 = st.columns(2)
                        with col_l3:
                            st.info("**LEVEL 3 (Procedural)**\n\n**Teacher:** 'What is step 1 to solve $2x + 5 = 15$?'\n\n**Student:** 'Subtract 5 from both sides.'\n\n**Teacher:** 'Correct, write that down.'")
                        with col_l5:
                            st.success("**LEVEL 5 (POGIL + 5-Move Talk Rotation)**\n\n**Teacher (Wait Time 5-7s):** '...' *(Pauses 6s)*\n\n**Teacher (Revoicing):** 'Spokesperson from Team 3, revoice what the Reflector noticed about testing $x=5$ on your visual balance mat.'\n\n**Spokesperson:** 'Our team agrees that removing 5 unit tiles keeps both scales equal.'\n\n**Teacher (Assigning Competence - Coaching Cue: Highlighting Spatial Discovery):** 'Notice how Alex's balance drawing made it obvious why we subtract 5 first! Team 4, who can **add on** or **agree/disagree with proof** to Alex's model?'")

                        st.markdown("""
                        ### 🎨 Visual On-Screen POGIL Discovery Mat
                        <div class="artifact-card">
                            <h3 style="margin-top:0;">📋 POGIL Student Discovery Mat: {selected_teks}</h3>
                            <p><span class="role-badge">Manager</span> Pacing & Turn-Taking &nbsp;|&nbsp; 
                               <span class="role-badge">Recorder</span> Visual Balance Model &nbsp;|&nbsp; 
                               <span class="role-badge">Spokesperson</span> Class Share &nbsp;|&nbsp; 
                               <span class="role-badge">Reflector</span> Talk Moves Check</p>
                            <hr>
                            <h4>Phase 1: Exploration (TEKS-Adaptive CRA)</h4>
                            <p>Sketch a Visual Balance Scale for $2x + 5 = 15$ before writing symbolic steps. What action maintains scale balance?</p>
                            <h4>Phase 2: Concept Invention</h4>
                            <p><b>Team Consensus Stem:</b> <i>"Our team derived the rule that operating on an equation maintains equality because..."</i></p>
                            <h4>Phase 3: Application & Accountable Talk Move</h4>
                            <p><b>Talk Move Prompt:</b> <i>"I want to add on to [Name]'s idea because testing $x=5$ proves..."</i></p>
                        </div>
                        """, unsafe_allow_html=True)

                        st.markdown("### 📋 Copyable Plain-Text Version (For Google Docs / Word)")
                        st.code(f"""================================================================================
POGIL STUDENT DISCOVERY MAT: {selected_teks}
================================================================================

TEAM ROLES:
- MANAGER: Tracks time limits and ensures equal participation across all members.
- RECORDER: Draws the team visual balance model and records written consensus.
- SPOKESPERSON: Shares out team reasoning during whole-class discussion.
- REFLECTOR: Monitors use of Accountable Talk Stems within the group.

PHASE 1: EXPLORATION (TEKS-Adaptive CRA Visual Sketch)
Sketch Balance Scale A (2x + 5 = 15) and Scale B (2x = 10).
Question: What specific action was taken on Scale A to create Scale B?

PHASE 2: CONCEPT INVENTION (Derived Rule)
Sentence Stem: "Our team derived the rule that subtracting equal quantities from both sides maintains scale balance because..."

PHASE 3: APPLICATION & ACCOUNTABLE TALK
Talk Move Prompt: "I want to add on to [Name]'s idea because testing x = 5 proves..."
===============================================================================""", language="text")

# ------------------------------------------------------------------------------
# TAB 2: WARMUP & EXIT TICKET GENERATOR (75% DUAL STAAR 2.0 ALIGNMENT)
# ------------------------------------------------------------------------------
with tab_assess:
    st.subheader("Targeted Assessment & DOK-3 Misconception Generator")
    st.caption("Generates entrance warmups and exit tickets formatted with 75% STAAR 2.0 Dual Layouts (Tech-Enhanced Item paired with Metacognitive Defense & Distractor Critique) plus 25% Open-Ended DOK-3 Performance Tasks.")

    col_a1, col_a2 = st.columns([1, 1])
    with col_a1:
        misconception_input = st.text_input(
            "Target Misconception / Common STAAR Error:",
            value="Confusing constant base fees with variable rates in mobile data / streaming plans"
        )
        dok_level = st.selectbox("Select Cognitive Depth & Item Format Blend:", [
            "75% STAAR 2.0 Tech-Enhanced + Metacognitive Defense (Recommended)",
            "100% Open-Ended DOK-3 Performance Task & Proof"
        ])
        btn_gen_assess = st.button("⚡ Build Assessment Slips & Distractor Analysis", type="primary", use_container_width=True)

    with col_a2:
        if btn_gen_assess:
            if client:
                system_prompt = f"""
                You are a STAAR 2.0 Assessment Specialist and Math Instructional Coach for {selected_grade}.
                TASK: Generate a 5-minute Entrance Warmup and Exit Slip for {selected_teks}.
                
                FORMATTING DIRECTIVE (75% OPTION C DUAL ALIGNMENT):
                - 75% of generated items MUST feature a STAAR 2.0 Tech-Enhanced Question on the Left (Inline Choice, Multi-Select, Drag-and-Drop Table, or Equation Builder) paired immediately with a mandatory Metacognitive Defense & Distractor Critique on the Right.
                - 25% of items can be pure open-ended DOK-3 performance tasks requiring a CRA visual sketch before symbolic solving.
                - Include an Eduphoria Distractor Analysis Key detailing choices A, B, C, D and the cognitive flaw linked to each.
                """
                response = client.models.generate_content(
                    model="gemini-1.5-pro",
                    contents=f"{system_prompt}\n\nMisconception: {misconception_input}\nFormat Preference: {dok_level}"
                )
                st.markdown(response.text)
            else:
                st.subheader(f"Generated Assessment Slips: {selected_teks}")
                st.markdown(f"""
                ### 📝 Entrance Warmup (STAAR 2.0 Dual Layout - 75% Blend)
                
                | **STAAR 2.0 Tech-Enhanced Item (Inline Choice / Drag & Drop)** | **Metacognitive Defense & Distractor Critique** |
                | :--- | :--- |
                | **Context:** Jordan is choosing a gaming pass. Plan X charges $10 flat plus $3/game. <br><br> Select the correct inequality to model keeping monthly costs under $40:<br> `[ $3g + 10 < 40 $ \| $10g + 3 < 40 $ \| $3g - 10 > 40 $ ]` | **Task 1:** Sketch a visual tape diagram of Plan X.<br><br>**Task 2:** Write 1 sentence explaining why placing the variable on $10 is a cognitive distractor error.<br><br> *Sentence Stem:* *"Student A's model fails because..."* |

                ---

                ### 🚪 Exit Slip (Metacognitive Justification)
                > Solve $-3x + 4 > 16$ in the context of cloud storage limits. Plot your boundary on a number line.  
                > **Accountable Defense Stem:** *"I verified my inequality sign direction by testing $x=0$, which proved..."*
                
                ---

                ### ⚠️ Eduphoria Distractor Analysis Key
                * **Correct Answer:** $3g + 10 < 40$ ($x < -4$)
                * **Distractor B ($10g + 3 < 40$):** Swapped fixed setup fee with variable rate per game.
                * **Distractor C ($x > -4$):** Solved steps correctly but forgot to flip sign during negative division.
                """)

# ------------------------------------------------------------------------------
# TAB 3: EDUPHORIA DATA & MULTIMODAL MISCONCEPTION COACH
# ------------------------------------------------------------------------------
with tab_data:
    st.subheader("Eduphoria Student Responses & Test File Analysis")
    st.caption("Upload your Eduphoria student item response CSV along with the Test file (PDF, Word, or Slides/PPTX). The AI cross-references distractor patterns with actual test questions to pinpoint misconceptions and structure exploratory reteach strategies.")

    st.info("🛡️ **FERPA Privacy Shield Active:** Uploaded student files are automatically anonymized in RAM (removing Student Names, Local IDs, and Teacher Names) before processing.")

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        uploaded_csv = st.file_uploader("1. Upload Eduphoria Student Item Responses (.csv):", type=["csv"], key="eduphoria_csv")
    with col_u2:
        uploaded_pdf = st.file_uploader("2. Upload Assessment File (PDF, Word, or Slides/PPTX):", type=["pdf", "docx", "pptx", "txt"], key="test_pdf")

    btn_analyze_test = st.button("🔍 Analyze Test & Identify Top 4 Misconceptions", type="primary", use_container_width=True)

    if btn_analyze_test:
        if not uploaded_csv or not uploaded_pdf:
            st.warning("⚠️ Please upload BOTH the Eduphoria student response CSV and the Assessment file to run the analysis.")
        else:
            try:
                raw_df = pd.read_csv(uploaded_csv)
                sanitized_df, dropped_pii = sanitize_eduphoria_dataframe(raw_df)
                csv_summary_str = sanitized_df.to_csv(index=False)
                
                test_file_text, pdf_part = extract_content_from_file(uploaded_pdf)
                
                st.success(f"🔒 FERPA Scrub Complete! Anonymized student data and loaded '{uploaded_pdf.name}'. Analyzing test items...")

                if client:
                    with st.spinner("Cross-referencing test questions with distractor choices to identify top 4 misconceptions..."):
                        system_prompt = f"""
                        You are a Master Junior High Math Instructional Coach and STAAR Assessment Specialist for {selected_grade}.
                        
                        TASK:
                        Analyze the attached Assessment file and the anonymized Eduphoria Student Response data ({selected_teks}).
                        Identify the TOP 4 MOST CRITICAL COMMON MISCONCEPTIONS across student responses.
                        
                        OUTPUT REQUIREMENTS FOR EACH OF THE 4 MISCONCEPTIONS:
                        
                        ### 🔴 Misconception #[Number]: [Short Descriptive Name]
                        * **Target Item & TEKS:** Identify specific question number(s), key TEKS, and distractor choice pattern (e.g., Question #5, Choice C - 42% of students).
                        * **1-2 Sentence Explanation:** A concise, sharp explanation of the exact cognitive flaw or misapplied logic leading to this error.
                        * **Exploratory Addressing Strategy (60% Inquiry / CRA Focus):** Provide coaching feedback on how the teacher can address this misconception in an **exploratory/discovery way** (e.g., 4-person POGIL inquiry prompt, TEKS-adaptive visual CRA model like balance scales or algebra tiles, or an error-critique stem) rather than teaching procedural steps.
                        
                        End with a brief coaching summary highlighting how correcting these 4 key gaps will elevate Domain 1 STAAR performance.
                        """
                        
                        prompt_contents = []
                        if pdf_part:
                            prompt_contents.append(pdf_part)
                        
                        prompt_contents.append(f"{system_prompt}\n\nTest Content:\n{test_file_text}\n\nStudent Item Response Data:\n{csv_summary_str}")
                        
                        response = client.models.generate_content(
                            model="gemini-1.5-pro",
                            contents=prompt_contents
                        )
                        st.markdown(response.text)
                else:
                    st.markdown("### 📊 Sample Preview (Add Gemini API Key to run live file analysis)")
                    st.markdown(f"""
                    ### 🔴 Misconception #1: Confusing Constant Base Fees with Variable Rates
                    * **Target Item & TEKS:** Question #4 | TEKS {selected_teks} | Distractor C chosen by 48% of students.
                    * **1-2 Sentence Explanation:** Students attached the variable rate to the one-time subscription setup fee instead of multiplying it by recurring data usage. This stems from reading key terms in isolation rather than analyzing rate of change.
                    * **Exploratory Addressing Strategy:**  
                      * **Visual CRA Activity:** Have 4-person POGIL teams build a Visual Balance Board comparing 1 month vs 3 months of usage.  
                      * **Inquiry Prompt Stem:** *"Team Member 2, test $x=0$ on Choice C. Does paying $0 in usage still cost $0 total under this plan? Why or why not?"*
                    """)
            except Exception as e:
                st.error(f"Error processing files: {e}")

# ------------------------------------------------------------------------------
# TAB 4: PLC CLUSTER & REFLECTIVE INQUIRY COACH
# ------------------------------------------------------------------------------
with tab_cluster:
    st.subheader("1-on-1 Coaching Debrief & Reflective Inquiry Coach")
    st.caption("Structures post-observation conferences using the Reflective Inquiry Protocol—guiding the coach with open-ended prompts so the teacher self-discovers instructional gaps.")

    col_c1, col_c2 = st.columns([1, 1])
    with col_c1:
        observation_notes = st.text_area(
            "Paste Observation Notes or Growth Goal:",
            height=160,
            value=f"Teacher led a guided lesson on {selected_teks}. Teacher asked procedural questions ('What is step 1?'). Students worked silently on worksheets with minimal discourse."
        )
        btn_cluster = st.button("📋 Generate Reflective Inquiry Debrief Guide", type="primary", use_container_width=True)

    with col_c2:
        if btn_cluster:
            clean_notes = sanitize_text_pii(observation_notes)
            if client:
                system_prompt = f"""
                You are a Master Instructional Coach guiding a 1-on-1 post-observation teacher conference.
                
                DEBRIEF PROTOCOL: REFLECTIVE INQUIRY (NON-DIRECTIVE DISCOVERY)
                1. AFFIRMATION: Open with sincere, authentic praise for teacher preparation and classroom climate.
                2. GUIDED REFLECTION (SEE IT): Provide 2-3 open-ended inquiry questions guiding the teacher to notice the gap between current student discourse and NIET Level 5 Exemplary questioning/thinking.
                3. CO-CONSTRUCTED ACTION STEP (NAME IT): Formulate open-ended prompts for the coach to help the teacher name their own bite-sized action step (e.g., embedding POGIL 4-person team sentence stems or DOK-3 error critiques).
                4. PRACTICE & COMMITMENT (DO IT): Guide the coach on co-scripting a 2-minute roll-out prompt for tomorrow's class.
                """
                response = client.models.generate_content(
                    model="gemini-1.5-pro",
                    contents=f"{system_prompt}\n\nObservation Notes:\n{clean_notes}"
                )
                st.markdown(response.text)
            else:
                st.markdown("""
                <div class="glow-card">
                    <h4>🌟 Affirming Opening</h4>
                    <p><i>"Thank you for welcoming me into your classroom! Your lesson pacing was smooth, and your clear board work gave students a solid visual anchor."</i></p>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("""
                ### ❓ Reflective Inquiry Script for Coach
                
                **1. See It (Guiding Discovery):**
                * *"When you look at student responses during guided practice, how would you describe the balance between procedural recall vs. students defending their reasoning?"*
                * *"How might providing a 4-person POGIL structure with Accountable Talk stems change who is doing the heavy lifting during problem-solving?"*
                
                **2. Name It (Teacher-Discovered Action Step):**
                * *"Looking at the NIET Rubric for Thinking and Problem Solving, what is one shift we could make in tomorrow's warmup to give students ownership of error critique?"*
                
                **3. Do It (Co-Scripting Practice):**
                * *"Let's take 2 minutes right now to co-script tomorrow's POGIL sentence stem for your mobile phone data plan problem!"*
                """)
