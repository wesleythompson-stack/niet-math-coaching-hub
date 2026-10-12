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
    .chat-container {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1rem;
        margin-top: 1.5rem;
    }
</style>
''', unsafe_allow_html=True)

# ==============================================================================
# 2. SESSION STATE INITIALIZATION FOR INTERACTIVE REASONING & CHAT
# ==============================================================================
if "elevator_output" not in st.session_state:
    st.session_state.elevator_output = None
if "elevator_history" not in st.session_state:
    st.session_state.elevator_history = []

if "assess_output" not in st.session_state:
    st.session_state.assess_output = None
if "assess_history" not in st.session_state:
    st.session_state.assess_history = []

if "chat_history" not in st.session_state:
    st.session_state.chat_history = [
        {
            "role": "assistant",
            "content": "👋 Hi! I am your NIET Math Coaching Assistant. Ask me to explain reasoning, correct my math logic, or generate alternate materials!"
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
# TAB 1: NIET LESSON ELEVATOR (INTERACTIVE REASONING & REGENERATION CHAT)
# ------------------------------------------------------------------------------
with tab_elevator:
    st.subheader("Elevate Lesson Draft or File Upload to NIET Level 5")
    st.caption("Upload a lesson file or paste your draft to receive concrete feedback, specific rephrased prompts, visual models, and interactive AI coaching.")

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
            placeholder=f"Example for {selected_grade}: Objective is to solve two-step inequalities or linear equations."
        )

        btn_elevate = st.button("🚀 Generate Level 5 NIET Coaching Brief", type="primary", use_container_width=True)

        if btn_elevate:
            if not lesson_text_input and not uploaded_lesson_file:
                st.warning("⚠️ Please paste lesson text or upload a document to generate feedback.")
            else:
                clean_text = sanitize_text_pii(lesson_text_input)
                file_text_content = ""
                pdf_part = None
                
                if uploaded_lesson_file:
                    file_text_content, pdf_part = extract_content_from_file(uploaded_lesson_file)
                
                combined_lesson_content = f"{clean_text}\n\n{file_text_content}".strip()
                
                with st.spinner("Analyzing lesson against NIET Level 5 Indicators..."):
                    if client:
                        prompt_contents = []
                        if pdf_part:
                            prompt_contents.append(pdf_part)
                        
                        sec_clause = f"and Secondary Focus: '{secondary_indicator}'" if secondary_indicator != "None" else ""
                        p1 = f"You are a Master Junior High Math Instructional Coach for {selected_grade} specializing in NIET rubrics."
                        p2 = f"PRIMARY NIET INDICATOR FOCUS: '{primary_indicator}' {sec_clause}."
                        p3 = f"TARGET TEKS: {selected_teks} ({selected_grade})."
                        
                        p_rules = (
                            "YOU MUST PROVIDE A FULL 5-SECTION COACHING BRIEF WITH SPECIFIC EXAMPLES:\n"
                            "### 🌟 Section 1: Authentic Affirming Glows (With Evidence)\n"
                            "### 🎯 Section 2: NIET Level 5 Coaching Grows & Concrete BEFORE vs AFTER Revisions\n"
                            "### ❓ Section 3: Targeted Clarification & Reflection Questions for the Teacher\n"
                            "### 🗣️ Section 4: Verbatim Complex Instruction & Accountable Talk Script\n"
                            "### 📜 Section 5: Elevated Student Discovery Mat (POGIL Activity)\n"
                        )
                        system_prompt = f"{p1}\n{p2}\n{p3}\n\n{p_rules}"
                        prompt_contents.append(f"{system_prompt}\n\nSubmitted Content:\n{combined_lesson_content}")
                        
                        response = client.models.generate_content(
                            model="gemini-1.5-pro",
                            contents=prompt_contents
                        )
                        st.session_state.elevator_output = response.text
                        st.session_state.elevator_history = []
                    else:
                        st.session_state.elevator_output = f"""
                        ### 🌟 Section 1: Authentic Affirming Glows
                        * **Clear Focus:** Targets **{selected_teks}** using real-world scenarios.
                        
                        ### 🎯 Section 2: NIET Level 5 Coaching Grows ({primary_indicator})
                        * **Original Prompt (Before):** *"Solve $3x + 10 < 40$."*
                        * **Elevated Version (After):** *"Model a balance scale comparing USD $10$ fixed cost with $3x$ recurring downloads."*

                        ### ❓ Section 3: Reflection Questions for the Teacher
                        1. *"How will student pairs divide roles during the visual discovery phase?"*
                        2. *"What scaffolding strategy will be used if students confuse setup fees with unit rates?"*

                        ### 🗣️ Section 4: Verbatim Accountable Talk Script
                        * **Teacher Cue:** *"Maya, can you rephrase Jordan's strategy in your own words?"*
                        """
                        st.session_state.elevator_history = []

    with col_e2:
        if st.session_state.elevator_output:
            st.markdown(st.session_state.elevator_output)
            
            st.divider()
            st.subheader("💡 Interactive AI Dialogue & Reasoning Coach")
            st.caption("You can ask the AI why it made specific choices, correct its reasoning, or request alternate lesson versions below:")

            # Render dialogue history
            for msg in st.session_state.elevator_history:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])

            # Interactive Input for Tab 1
            if elevator_feedback := st.chat_input("Ask why AI chose a strategy, correct its logic, or request alternative versions...", key="chat_tab1"):
                clean_fb = sanitize_text_pii(elevator_feedback)
                st.session_state.elevator_history.append({"role": "user", "content": clean_fb})
                
                with st.chat_message("user"):
                    st.markdown(clean_fb)

                with st.chat_message("assistant"):
                    if client:
                        with st.spinner("Processing feedback, reviewing reasoning, or regenerating material..."):
                            elevator_system_instruction = (
                                f"You are a collaborative NIET Math Instructional Specialist for {selected_grade} (TEKS Focus: {selected_teks}).\n"
                                f"You previously generated the following Level 5 Coaching Brief:\n"
                                f"--- START GENERATED BRIEF ---\n{st.session_state.elevator_output}\n--- END GENERATED BRIEF ---\n\n"
                                "YOUR MANDATORY INTERACTION DIRECTIVES:\n"
                                "1. EXPLAIN REASONING: If the user asks WHY you recommended a specific script, prompt, or CRA visual model, explain your pedagogical rationale clearly.\n"
                                "2. ACCEPT & INCORPORATE CORRECTIONS: If the user points out a flaw or corrects your math/TEKS logic, validate their correction, acknowledge the mistake, and update your logic.\n"
                                "3. REGENERATE ALTERNATIVES: If asked for alternative versions (e.g., accommodations for EL/ML learners, different grade levels, alternate POGIL mats), output the revised material or full updated brief in clean markdown."
                            )
                            
                            # Construct multi-turn context
                            conversation_context = f"{elevator_system_instruction}\n\nCONVERSATION HISTORY:\n"
                            for item in st.session_state.elevator_history:
                                role_label = "User" if item["role"] == "user" else "AI Coach"
                                conversation_context += f"{role_label}: {item['content']}\n"
                            
                            resp = client.models.generate_content(
                                model="gemini-1.5-pro",
                                contents=conversation_context
                            )
                            ai_fb_resp = resp.text
                            st.markdown(ai_fb_resp)
                            st.session_state.elevator_history.append({"role": "assistant", "content": ai_fb_resp})
                    else:
                        fallback_msg = "Connect your Gemini API key to interactively question reasoning or regenerate results."
                        st.markdown(fallback_msg)
                        st.session_state.elevator_history.append({"role": "assistant", "content": fallback_msg})

# ------------------------------------------------------------------------------
# TAB 2: WARMUP & EXIT TICKET (INTERACTIVE REASONING & REGENERATION CHAT)
# ------------------------------------------------------------------------------
with tab_assess:
    st.subheader("Targeted Assessment & Lesson Objective Generator")
    st.caption("Upload a lesson activity to extract its objective and generate an entrance launch warmup and objective mastery exit ticket.")

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

        if btn_gen_assess:
            if not uploaded_assess_file and not misconception_input.strip():
                st.warning("⚠️ Please upload a lesson file or enter an objective/misconception to continue.")
            else:
                clean_text = sanitize_text_pii(misconception_input)
                file_text_content = ""
                pdf_part = None
                
                if uploaded_assess_file:
                    file_text_content, pdf_part = extract_content_from_file(uploaded_assess_file)
                
                combined_assess_input = f"Manual Input: {clean_text}\n\nUploaded Content:\n{file_text_content}".strip()
                
                with st.spinner("Extracting lesson objective and generating Launch & Exit Ticket..."):
                    if client:
                        prompt_contents = []
                        if pdf_part:
                            prompt_contents.append(pdf_part)
                        
                        p1 = f"You are a STAAR 2.0 Assessment Specialist and Math Coach for {selected_grade}."
                        p2 = f"TARGET TEKS: {selected_teks} ({selected_grade})."
                        p_rules = (
                            "Generate:\n"
                            "1. Extracted Objective & Target Misconception\n"
                            "2. 5-min Warmup Lesson Launch (4 options: 1 correct, 2 procedural, 1 instant elimination via vocab/estimation)\n"
                            "3. 5-min Exit Ticket Assessment (4 options: 1 correct, 2 procedural, 1 instant elimination)\n"
                            "4. Eduphoria Distractor Analysis Key\n"
                        )
                        system_prompt = f"{p1}\n{p2}\n\n{p_rules}"
                        prompt_contents.append(f"{system_prompt}\n\nInput Content:\n{combined_assess_input}")
                        
                        response = client.models.generate_content(
                            model="gemini-1.5-pro",
                            contents=prompt_contents
                        )
                        st.session_state.assess_output = response.text
                        st.session_state.assess_history = []
                    else:
                        st.session_state.assess_output = f"""
                        ### 🎯 Extracted Lesson Objective & TEKS Focus
                        * **Objective:** Solve two-step inequalities representing mobile cloud pass subscriptions.
                        * **Misconception:** Confusing one-time setup fees with recurring rates.

                        ### 📝 Entrance Warmup: Lesson Launch
                        **Task:** Jordan sets up a cloud pass with a USD $10 fee plus USD $3 per download. Spending under USD $40:
                        * [ ] **Option A:** $3g + 10 < 40$
                        * [ ] **Option B:** $10g + 3 < 40$
                        * [ ] **Option C:** $3g + 10 > 40$
                        * [ ] **Option D:** $3g + 10 < 4$

                        ### 🚪 Exit Ticket: Objective Mastery Assessment
                        **Task:** Solve $-3x + 4 > 16$:
                        * [ ] **Option A:** $x < -4$
                        * [ ] **Option B:** $x > -4$
                        * [ ] **Option C:** $x < 4$
                        * [ ] **Option D:** $x < 400$
                        """
                        st.session_state.assess_history = []

    with col_a2:
        if st.session_state.assess_output:
            st.markdown(st.session_state.assess_output)
            
            st.divider()
            st.subheader("💡 Interactive Assessment Reasoning & Regeneration Coach")
            st.caption("Ask why a distractor was chosen, correct option errors, or ask for alternative question formats below:")

            # Render dialogue history
            for msg in st.session_state.assess_history:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])

            # Interactive Input for Tab 2
            if assess_feedback := st.chat_input("Ask why distractor B was selected, correct an answer option, or request Form B...", key="chat_tab2"):
                clean_fb = sanitize_text_pii(assess_feedback)
                st.session_state.assess_history.append({"role": "user", "content": clean_fb})
                
                with st.chat_message("user"):
                    st.markdown(clean_fb)

                with st.chat_message("assistant"):
                    if client:
                        with st.spinner("Reviewing assessment rationale or generating modified questions..."):
                            assess_system_instruction = (
                                f"You are a STAAR 2.0 Assessment Specialist & Math Coach for {selected_grade} (TEKS Focus: {selected_teks}).\n"
                                f"You previously generated the following Lesson Launch & Exit Ticket:\n"
                                f"--- START GENERATED ASSESSMENT ---\n{st.session_state.assess_output}\n--- END GENERATED ASSESSMENT ---\n\n"
                                "YOUR MANDATORY INTERACTION DIRECTIVES:\n"
                                "1. EXPLAIN REASONING: If the user asks why a specific distractor, option, or metacognitive task was chosen, explain the underlying student misconception or STAAR error pattern.\n"
                                "2. ACCEPT & INCORPORATE CORRECTIONS: If the user states a distractor is incorrect, mathematically flawed, or too obvious, validate their critique and fix it immediately.\n"
                                "3. REGENERATE ALTERNATIVES: If asked for alternate forms (Form B), different real-world scenarios, or modified difficulty levels (e.g., fractions/decimals), regenerate full updated questions maintaining the mandatory 4-option multiple choice structure and analysis key."
                            )
                            
                            # Construct multi-turn context
                            conversation_context = f"{assess_system_instruction}\n\nCONVERSATION HISTORY:\n"
                            for item in st.session_state.assess_history:
                                role_label = "User" if item["role"] == "user" else "AI Assessment Coach"
                                conversation_context += f"{role_label}: {item['content']}\n"

                            resp = client.models.generate_content(
                                model="gemini-1.5-pro",
                                contents=conversation_context
                            )
                            ai_assess_resp = resp.text
                            st.markdown(ai_assess_resp)
                            st.session_state.assess_history.append({"role": "assistant", "content": ai_assess_resp})
                    else:
                        fallback_msg = "Connect your Gemini API key to interactively question assessment distractors or regenerate alternate items."
                        st.markdown(fallback_msg)
                        st.session_state.assess_history.append({"role": "assistant", "content": fallback_msg})

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
                        
                        prompt_contents.append(f"{system_prompt}\n\nTest Content:\n{test_file_text}\n\nStudent Data:\n{csv_summary_str}")
                        
                        response = client.models.generate_content(
                            model="gemini-1.5-pro",
                            contents=prompt_contents
                        )
                        st.markdown(response.text)
                else:
                    st.markdown("### 📊 Sample Misconception Analysis")
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
# TAB 5: GLOBAL CHAT ASSISTANT
# ------------------------------------------------------------------------------
with tab_chat:
    st.subheader("💬 Global Coaching Assistant Chat")
    st.caption("Ask general pedagogical questions, explore NIET indicators, or seek advice across topics.")

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

    for message in st.session_state.chat_history:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    if chat_prompt := st.chat_input("Ask any math coaching question...", key="chat_global"):
        clean_chat_prompt = sanitize_text_pii(chat_prompt)
        st.session_state.chat_history.append({"role": "user", "content": clean_chat_prompt})
        
        with st.chat_message("user"):
            st.markdown(clean_chat_prompt)

        with st.chat_message("assistant"):
            if client:
                with st.spinner("Analyzing request..."):
                    system_context = (
                        f"You are a Master Junior High Math Instructional Coach for {selected_grade}.\n"
                        f"ACTIVE COURSE: {selected_grade}\n"
                        f"ACTIVE TEKS FOCUS: {selected_teks}\n\n"
                        "DIRECTIVES:\n"
                        "1. Assist teachers and coaches with clarifying materials, scaffolding, accommodations, or adjusting questions.\n"
                        "2. Respond with actionable instructional wisdom.\n"
                        "3. Use LaTeX formatting for formal math expressions ($2x + 5 = 15$).\n"
                        "4. Do NOT use raw HTML tags."
                    )
                    
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
                fallback_response = "Connect your Gemini API key to chat with the coaching assistant."
                st.markdown(fallback_response)
                st.session_state.chat_history.append({"role": "assistant", "content": fallback_response})
