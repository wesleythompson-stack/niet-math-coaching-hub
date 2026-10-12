import re
import os
import io
import pandas as pd
import streamlit as st
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
        padding: 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
    }
    .artifact-card {
        background-color: #ffffff;
        border: 2px dashed #94a3b8;
        padding: 1rem;
        border-radius: 8px;
        margin-top: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. FERPA SANITIZATION ENGINE (OPTION A COMPLIANCE)
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
    <p style="margin:0; opacity:0.9;">Tailored for <b>{selected_grade}</b> • 60% Conceptual Discovery + 40% CRA Modeling • NIET Level 5 Reflective Framework</p>
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
# TAB 1: NIET LESSON ELEVATOR
# ------------------------------------------------------------------------------
with tab_elevator:
    st.subheader("Elevate Lesson Draft to NIET Level 5")
    st.caption("Transforms standard lesson drafts using 60% Conceptual Discovery, 40% CRA visual modeling, DOK-3 cognitive escalation, and teen-relevant contexts.")

    col1, col2 = st.columns([1, 1])
    with col1:
        niet_indicator = st.selectbox(
            "Primary NIET Indicator Focus:",
            ["Questioning (Varied, Sequenced & Student-Generated Inquiry)", 
             "Thinking & Problem Solving (Analytical, Practical, Creative & Metacognition)", 
             "Academic Feedback (Frequent, High-Quality & Peer Critique)", 
             "Grouping Students (4-Person Kagan Teams & Individual Accountability)"]
        )
        
        lesson_input = st.text_area(
            "Paste lesson draft / objective here:",
            height=180,
            placeholder=f"Example for {selected_grade}: Students will set up and solve equations. Teacher presents 2 examples on board. Students work in pairs on a worksheet."
        )

        btn_elevate = st.button("🚀 Generate Level 5 Coaching Brief", type="primary", use_container_width=True)

    with col2:
        if btn_elevate:
            clean_lesson = sanitize_text_pii(lesson_input)
            with st.spinner("Analyzing lesson with Master Coach lens..."):
                if client:
                    system_prompt = f"""
                    You are a Master Junior High Math Instructional Coach specializing in NIET Rubric alignment for {selected_grade}.
                    
                    MANDATORY COACHING PHILOSOPHY & PEDAGOGICAL DIRECTIVES:
                    1. EMPATHETIC TONE & GLOWS: Always start with enthusiastic, authentic 'Glows' acknowledging the teacher's intentionality and standards alignment.
                    2. NIET LEVEL 5 DESCRIPTOR ALIGNMENT: Explicitly reference exact NIET Rubric Level 5 performance descriptors for '{niet_indicator}'. Point out what shifts the lesson from Level 3 (Proficient) to Level 5 (Exemplary).
                    3. DOK-3 ESCALATION & METACOGNITION: Convert any low-level procedural tasks (DOK 1/2) into high-level DOK-3 cognitive tasks. Force students to evaluate their own thinking, defend solutions, and critique flawed reasoning.
                    4. TEEN-RELEVANT CONTEXTS: Frame mathematical situations around real 7th/9th grade scenarios (e.g., mobile phone data caps, cloud storage limits, gaming pass tiers, streaming algorithm metrics, sports analytics).
                    5. 4-PERSON KAGAN GROUP DISCOURSE: Structure team discourse using 4-person mixed-ability Kagan roles (RallyCoach, RoundRobin) with structured academic sentence stems leading to whole-class share-outs.
                    6. 60% DISCOVERY / 40% CRA MODELING: Ensure inquiry and visual/concrete tools (algebra tiles, visual balance scales, area models) precede procedural rules.
                    
                    OUTPUT FORMAT REQUIRED:
                    ### 🌟 Authentic Affirming Glows
                    [Affirm teacher intent, clarity, and standards alignment]
                    
                    ### 🎯 NIET Level 5 Indicator Breakdown ({niet_indicator})
                    [Explicitly cite NIET Level 5 rubric descriptors and map out the exact shift required]
                    
                    ### 💬 Classroom Dialogue Script (Level 3 Procedural vs Level 5 Exemplary)
                    - **Level 3 (Teacher-Led Procedural Recall):** Standard dialogue.
                    - **Level 5 (Student-Led Metacognitive Discovery):** Dialogue featuring teen-relevant context, 4-person Kagan group roles, sentence stems, and DOK-3 defense.
                    
                    ### 🖨️ Printable Student Scaffold Artifact
                    [Create a text-based printable graphic organizer / Kagan team discovery mat incorporating visual CRA models and defense stems]
                    """
                    
                    response = client.models.generate_content(
                        model="gemini-1.5-pro",
                        contents=f"{system_prompt}\n\nTarget TEKS: {selected_teks}\nIndicator: {niet_indicator}\nLesson Draft:\n{clean_lesson}"
                    )
                    st.markdown(response.text)
                else:
                    st.markdown(f"""
                    <div class="glow-card">
                        <h4>🌟 Authentic Affirming Glows</h4>
                        <p>Wonderful foundation! Your draft for <b>{selected_teks}</b> establishes clear mathematical objectives and structures guided practice for students.</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.markdown(f"### 🎯 NIET Level 5 Descriptor Shift ({niet_indicator})")
                    st.write("**NIET Exemplary Target:** *'Questions consistently engage students in critical thinking, require students to draw connections, and prompt metacognitive self-monitoring.'*")
                    
                    st.markdown(f"### 💬 Level 3 vs Level 5 Dialogue Script ({selected_grade})")
                    col_l3, col_l5 = st.columns(2)
                    with col_l3:
                        st.info("**LEVEL 3 (Procedural Recall)**\n\n**Teacher:** 'To calculate monthly cloud storage cost $C = 0.15g + 5$, what do we plug in for 50 GB?'\n\n**Student:** 'Multiply 0.15 by 50 and add 5 to get $12.50.'\n\n**Teacher:** 'Correct, move to problem #2.'")
                    with col_l5:
                        st.success(f"**LEVEL 5 (DOK-3 Metacognition & Kagan Discourse)**\n\n**Teacher:** 'Team Member 1, model the storage plan $C = 0.15g + 5$ on your visual balance mat. Member 2, critique Student B's claim that doubling storage from 50 GB to 100 GB doubles the total bill. Members 3 & 4, prepare your team defense.'\n\n**Sentence Stem:** *'Student B's logic fails because the $5 base fee remains constant, meaning...'*\n\n**Student 2 (Team Defense):** '50 GB costs $12.50, but 100 GB costs $20.00, not $25.00! The constant $5 fee prevents it from being a direct proportion.'")

                    st.markdown("""
                    <div class="artifact-card">
                        <h4>🖨️ Printable Student Artifact: Teen Context Discovery Mat</h4>
                        <p><b>Context:</b> Gaming Subscription vs. Per-Game Purchase Model</p>
                        <hr>
                        <p><b>Step 1 (Visual CRA Model):</b> Sketch the monthly cost scale for Plan A ($15 base + $2/game) vs Plan B ($30 unlimited).<br>
                        <b>Step 2 (Metacognitive Defense Stem):</b> <i>"Our team proved Plan A becomes more expensive than Plan B when a gamer downloads more than _____ games because..."</i><br>
                        <b>Step 3 (Error Critique):</b> Identify why Distractor B incorrectly calculated the break-even point.</p>
                    </div>
                    """, unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# TAB 2: WARMUP & EXIT TICKET GENERATOR
# ------------------------------------------------------------------------------
with tab_assess:
    st.subheader("Targeted Assessment & DOK-3 Misconception Generator")
    st.caption("Generates warmups and exit tickets targeting high-frequency STAAR distractors, utilizing teen contexts and requiring student error critique and justification.")

    col_a1, col_a2 = st.columns([1, 1])
    with col_a1:
        misconception_input = st.text_input(
            "Target Misconception / Common STAAR Error:",
            value="Confusing constant base fees with variable rates in mobile data / streaming plans"
        )
        dok_level = st.selectbox("Select Cognitive Depth:", ["DOK 3 - Error Critique & Flawed Logic Defense", "DOK 3 - Multi-Step Modeling & Proof"])
        btn_gen_assess = st.button("⚡ Build Assessment Slips & Distractor Analysis", type="primary", use_container_width=True)

    with col_a2:
        if btn_gen_assess:
            if client:
                system_prompt = f"""
                You are a STAAR Math Assessment Specialist for {selected_grade}.
                TASK: Generate a 5-minute Entrance Warmup and Exit Slip for {selected_teks}.
                REQUIREMENTS:
                1. TEEN CONTEXT: Use scenarios like mobile gaming passes, cloud storage, streaming video data, or athletic stats.
                2. DOK-3 ESCALATION & ERROR DEFENSE: Present flawed student solutions. Ask students to locate the misconception, correct it visually, and write a metacognitive defense.
                3. EDUPHORIA DISTRACTOR KEY: Detail choices A, B, C, D and explain the exact cognitive mistake associated with each distractor.
                """
                response = client.models.generate_content(
                    model="gemini-1.5-pro",
                    contents=f"{system_prompt}\n\nMisconception: {misconception_input}\nDOK Target: {dok_level}"
                )
                st.markdown(response.text)
            else:
                st.subheader(f"Generated Assessment Slips: {selected_teks}")
                st.markdown(f"""
                **WARMUP (Error Critique & Defense - DOK 3)**
                > *Context:* Jordan is comparing gaming downloads. Plan X charges a $10 flat fee plus $3 per game. Student A wrote the inequality $10g + 3 < 40$ to find how many games he can download for under $40.  
                > **Task:** Draw a visual tape diagram of Plan X. Explain why Student A's inequality is incorrect and rewrite the correct inequality.  
                > **Sentence Stem:** *"Student A attached the variable to the wrong quantity because..."*
                
                **EXIT SLIP (Metacognitive Justification - DOK 3)**
                > Solve $-3x + 4 > 16$ in the context of streaming data limits. Plot your solution on a visual number line.  
                > **Sentence Stem:** *"I verified my boundary arrow direction by testing $x=0$, which proved..."*
                
                **⚠️ Eduphoria Distractor Analysis Key:**
                * **Correct Answer:** $x < -4$
                * **Distractor B ($x > -4$):** Solved algebraic steps correctly but forgot to flip the inequality sign when dividing by $-3$.
                * **Distractor C ($x < -6$):** Sign error during subtraction ($16 - 4 = -12$).
                """)

# ------------------------------------------------------------------------------
# TAB 3: EDUPHORIA DATA & MULTIMODAL MISCONCEPTION COACH
# ------------------------------------------------------------------------------
with tab_data:
    st.subheader("Eduphoria Student Responses & Test PDF Analysis")
    st.caption("Upload your Eduphoria student item response CSV along with the Test PDF. The AI cross-references distractor patterns with actual test questions to pinpoint misconceptions and structure exploratory reteach strategies.")

    st.info("🛡️ **FERPA Privacy Shield Active:** Uploaded student files are automatically anonymized in RAM (removing Student Names, Local IDs, and Teacher Names) before processing.")

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        uploaded_csv = st.file_uploader("1. Upload Eduphoria Student Item Responses (.csv):", type=["csv"], key="eduphoria_csv")
    with col_u2:
        uploaded_pdf = st.file_uploader("2. Upload Assessment PDF Copy (.pdf):", type=["pdf"], key="test_pdf")

    btn_analyze_test = st.button("🔍 Analyze Test & Identify Top 4 Misconceptions", type="primary", use_container_width=True)

    if btn_analyze_test:
        if not uploaded_csv or not uploaded_pdf:
            st.warning("⚠️ Please upload BOTH the Eduphoria student response CSV and the Assessment PDF to run the analysis.")
        else:
            try:
                # 1. Process & Sanitize CSV
                raw_df = pd.read_csv(uploaded_csv)
                sanitized_df, dropped_pii = sanitize_eduphoria_dataframe(raw_df)
                csv_summary_str = sanitized_df.to_csv(index=False)
                
                # 2. Process PDF Bytes for Gemini Multimodal Input
                pdf_bytes = uploaded_pdf.read()
                pdf_part = types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")
                
                st.success(f"🔒 FERPA Scrub Complete! Anonymized student data and loaded '{uploaded_pdf.name}'. Analyzing test items...")

                if client:
                    with st.spinner("Cross-referencing test questions with distractor choices to identify top 4 misconceptions..."):
                        system_prompt = f"""
                        You are a Master Junior High Math Instructional Coach and STAAR Assessment Specialist for {selected_grade}.
                        
                        TASK:
                        Analyze the attached Assessment PDF and the anonymized Eduphoria Student Response data ({selected_teks}).
                        Identify the TOP 4 MOST CRITICAL COMMON MISCONCEPTIONS across student responses.
                        
                        OUTPUT REQUIREMENTS FOR EACH OF THE 4 MISCONCEPTIONS:
                        
                        ### 🔴 Misconception #[Number]: [Short Descriptive Name]
                        * **Target Item & TEKS:** Identify the specific question number(s), key TEKS, and distractor choice pattern (e.g., Question #5, Choice C - 42% of students).
                        * **1-2 Sentence Explanation:** A concise, sharp explanation of the exact cognitive flaw or misapplied logic leading to this error.
                        * **Exploratory Addressing Strategy (60% Inquiry / CRA Focus):** Provide coaching feedback on how the teacher can address this misconception with students in an **exploratory/discovery way** (e.g., using a 4-person Kagan inquiry prompt, visual CRA model like balance mats or algebra tiles, or an error-critique stem) rather than just teaching the procedural step.
                        
                        End with a brief, encouraging coaching summary highlighting how correcting these 4 key gaps will elevate Domain 1 STAAR performance.
                        """
                        
                        response = client.models.generate_content(
                            model="gemini-1.5-pro",
                            contents=[
                                pdf_part,
                                f"{system_prompt}\n\nStudent Item Response Data:\n{csv_summary_str}"
                            ]
                        )
                        st.markdown(response.text)
                else:
                    st.markdown("### 📊 Sample Preview (Add Gemini API Key to run live PDF analysis)")
                    st.markdown(f"""
                    ### 🔴 Misconception #1: Confusing Constant Base Fees with Variable Rates
                    * **Target Item & TEKS:** Question #4 | TEKS {selected_teks} | Distractor C chosen by 48% of students.
                    * **1-2 Sentence Explanation:** Students attached the variable rate to the one-time subscription setup fee instead of multiplying it by recurring data usage. This stems from reading key terms in isolation rather than analyzing rate of change.
                    * **Exploratory Addressing Strategy:**  
                      * **Visual CRA Activity:** Have 4-person Kagan teams build a Visual Balance Board comparing 1 month vs 3 months of usage.  
                      * **Inquiry Prompt Stem:** *"Team Member 2, test $x=0$ on Choice C. Does paying $0 in usage still cost $0 total under this plan? Why or why not?"*

                    ---

                    ### 🔴 Misconception #2: Neglecting Inequality Sign Reversal During Division
                    * **Target Item & TEKS:** Question #8 | TEKS {selected_teks} | Distractor B chosen by 38% of students.
                    * **1-2 Sentence Explanation:** Students correctly applied inverse operations to isolate the variable, but treated the inequality symbol like an equals sign during negative division.
                    * **Exploratory Addressing Strategy:**  
                      * **Discovery Activity:** Ask teams to plot test values $x=5$ and $x=-5$ on a horizontal scale mat.  
                      * **Inquiry Prompt Stem:** *"Our team discovered that dividing by a negative number flips our arrow because testing $x=0$ proves..."*

                    ---

                    ### 🔴 Misconception #3: Incorrect Area Decomposition of Composite Figures
                    * **Target Item & TEKS:** Question #11 | TEKS {selected_teks} | Distractor A chosen by 35% of students.
                    * **1-2 Sentence Explanation:** Students added overlapping boundary lengths when computing total perimeter instead of decomposing the interior regions.
                    * **Exploratory Addressing Strategy:**  
                      * **Visual CRA Activity:** Provide color-coded grid overlays for students to physically cut and reassemble composite shapes before running calculations.

                    ---

                    ### 🔴 Misconception #4: Misinterpreting Direct Variation Scale Factors
                    * **Target Item & TEKS:** Question #14 | TEKS {selected_teks} | Distractor D chosen by 31% of students.
                    * **1-2 Sentence Explanation:** Students subtracted coordinate points ($y_2 - y_1$) rather than dividing to find the constant ratio $k = y/x$.
                    * **Exploratory Addressing Strategy:**  
                      * **Discovery Activity:** Have teams create a ratio table for 3 real-world scenarios and test if the ratio remains constant across all pairs.
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
                2. GUIDED REFLECTION (SEE IT): Provide 2-3 open-ended inquiry questions that guide the teacher to notice the gap between current student discourse and NIET Level 5 Exemplary questioning/thinking.
                3. CO-CONSTRUCTED ACTION STEP (NAME IT): Formulate open-ended prompts for the coach to help the teacher name their own bite-sized action step (e.g., embedding Kagan 4-person team sentence stems or DOK-3 error critiques).
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
                * *"How might providing a 4-person Kagan structure with sentence stems change who is doing the heavy lifting during problem-solving?"*
                
                **2. Name It (Teacher-Discovered Action Step):**
                * *"Looking at the NIET Rubric for Thinking and Problem Solving, what is one shift we could make in tomorrow's warmup to give students ownership of error critique?"*
                
                **3. Do It (Co-Scripting Practice):**
                * *"Let's take 2 minutes right now to co-script tomorrow's Kagan sentence stem for your mobile phone data plan problem!"*
                """)
