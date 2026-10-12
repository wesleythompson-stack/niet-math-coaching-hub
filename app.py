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
# 3. SIDEBAR CONFIGURATION & API KEY SETUP
# ==============================================================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/calculator.png", width=64)
    st.title("Coaching Settings")
    
    # Priority: Secrets Manager -> Sidebar Fallback
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        api_key = st.text_input("Enter Gemini API Key (Optional)", type="password")
    
    st.divider()
    
    selected_grade = st.radio(
        "Select Target Course:",
        ["7th Grade Math", "9th Grade Algebra 1"]
    )
    
    if selected_grade == "7th Grade Math":
        teks_presets = ["7.11A (Two-Step Inequalities)", "7.4A (Constant Rate of Change)", "7.9A (Area of Composite Figures)", "Custom TEKS"]
    else:
        teks_presets = ["A.5A (Linear Equations)", "A.2I (Writing Systems of Equations)", "A.3C (Graphing Linear Functions)", "Custom TEKS"]
        
    selected_teks = st.selectbox("Select TEKS Focus:", teks_presets)
    if selected_teks == "Custom TEKS":
        selected_teks = st.text_input("Enter Custom TEKS (e.g., 7.10B or A.12E):")

    st.divider()
    st.caption("🛡️ **FERPA Protected:** Eduphoria files are sanitized in RAM before API calls.")

client = genai.Client(api_key=api_key) if api_key else None

# ==============================================================================
# 4. HEADER
# ==============================================================================
st.markdown(f"""
<div class="main-header">
    <h1 style="margin:0;">📐 NIET Math Coaching Intelligence Hub</h1>
    <p style="margin:0; opacity:0.9;">Tailored for <b>{selected_grade}</b> • 60% Conceptual Discovery + 40% CRA Modeling • Eduphoria Aware Integrated</p>
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# 5. MAIN TABS
# ==============================================================================
tab_elevator, tab_assess, tab_data, tab_cluster = st.tabs([
    "✨ NIET Lesson Elevator",
    "📝 Warmup & Exit Ticket Generator",
    "📊 Eduphoria Data & STAAR Coach",
    "👥 PLC Cluster & Growth Coach"
])

# ------------------------------------------------------------------------------
# TAB 1: NIET LESSON ELEVATOR
# ------------------------------------------------------------------------------
with tab_elevator:
    st.subheader("Elevate Lesson Draft to NIET Level 5")
    st.caption("Transforms standard lesson drafts using 60% Conceptual Discovery, 40% CRA visual modeling, and 4-person mixed-ability Kagan group discourse.")

    col1, col2 = st.columns([1, 1])
    with col1:
        niet_indicator = st.selectbox(
            "Primary NIET Indicator Focus:",
            ["Questioning (Higher-Order & Wait Time)", 
             "Academic Feedback & Self-Correction", 
             "Thinking & Problem Solving (Analytical/Creative)", 
             "Grouping Students (4-Person Kagan Teams & Individual Accountability)"]
        )
        
        lesson_input = st.text_area(
            "Paste lesson draft / objective here:",
            height=180,
            placeholder=f"Example for {selected_grade}: Objective is to solve equations. Teacher demonstrates 3 examples on board. Students complete a worksheet."
        )

        btn_elevate = st.button("🚀 Generate Level 5 Coaching Brief", type="primary", use_container_width=True)

    with col2:
        if btn_elevate:
            clean_lesson = sanitize_text_pii(lesson_input)
            with st.spinner("Analyzing lesson with Master Coach lens..."):
                if client:
                    system_prompt = f"""
                    You are a Master Junior High Math Instructional Coach specializing in NIET Rubric alignment for {selected_grade}.
                    
                    NON-NEGOTIABLE PEDAGOGICAL RULES:
                    1. TONE: Empathetic, supportive, and affirming. ALWAYS lead with authentic, enthusiastic 'Glows' to celebrate teacher intent before introducing growth steps.
                    2. MATH RIGOR: 60% Conceptual Discovery (student inquiry, error critique) + 40% CRA Models (algebra tiles, area models, visual balance scales) BEFORE procedure.
                    3. GROUP DISCOURSE: Structure discourse around 4-person mixed-ability Kagan set teams (RallyCoach, RoundRobin) with structured sentence stems, leading to whole-class share-outs.
                    4. TEKS FOCUS: {selected_teks} ({selected_grade}).
                    
                    OUTPUT STRUCTURE REQUIRED:
                    ### 🌟 Authentic Affirming Glows
                    [Celebrate teacher intent, clarity, and standards alignment]
                    
                    ### 💬 Level 3 vs Level 5 Classroom Dialogue Script ({niet_indicator})
                    - **Level 3 (Procedural Recall):** Show standard teacher/student dialogue.
                    - **Level 5 ({selected_grade} Exemplary):** Show exact script featuring 4-person Kagan group talk, CRA visual models, and sentence stems.
                    
                    ### 🖨️ Printable Student Scaffold Artifact
                    [Create a text-based printable graphic organizer or student mat]
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
                        <p>Wonderful foundation! Your lesson objective for <b>{selected_teks}</b> is clearly aligned with grade-level rigor, and your planned guided practice shows a strong commitment to structured support!</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.markdown(f"### 💬 Level 3 vs Level 5 Script ({selected_grade})")
                    col_l3, col_l5 = st.columns(2)
                    with col_l3:
                        st.info("**LEVEL 3 (Procedural)**\n\n**Teacher:** 'What step do we take to solve $-3x + 4 > 16$?'\n\n**Student:** 'Subtract 4, then divide by -3 and flip the sign.'\n\n**Teacher:** 'Correct! Start problem #1.'")
                    with col_l5:
                        st.success(f"**LEVEL 5 (60% Discovery + CRA + 4-Person Kagan Teams)**\n\n**Teacher:** 'Team Member 1 & 2, open your Visual Balance Mat. Member 3, read the prompt. Member 4, record team consensus.'\n\n**Sentence Stem:** *'Our team tested $x=0$ and discovered...'*\n\n**Student 2 (Team Share):** 'When we tested 0, $-3(0)+4 > 16$ became $4 > 16$, which is false! That proves our arrow must point left.'")

                    st.markdown("""
                    <div class="artifact-card">
                        <h4>🖨️ Printable Student Artifact: Kagan Team Discovery Mat</h4>
                        <p><b>Target:</b> TEKS Verification & Proof Scale</p>
                        <hr>
                        <p><b>Step 1:</b> Write team consensus prediction.<br>
                        <b>Step 2 (Sentence Stem):</b> <i>"If we divide by a negative scale factor, the balance reverses because..."</i><br>
                        <b>Step 3:</b> Plot test values $x=0$ and $x=-5$ on your team balance board.</p>
                    </div>
                    """, unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# TAB 2: WARMUP & EXIT TICKET GENERATOR
# ------------------------------------------------------------------------------
with tab_assess:
    st.subheader("Targeted Assessment & Misconception Generator")
    st.caption("Generates entrance warmups and exit tickets targeting high-frequency STAAR Eduphoria distractors while requiring visual justification.")

    col_a1, col_a2 = st.columns([1, 1])
    with col_a1:
        misconception_input = st.text_input(
            "Target Misconception / Common Error:",
            value="Dropping negative sign when distributing or failing to reverse inequality sign"
        )
        dok_level = st.selectbox("Select DOK Target:", ["DOK 2 - Conceptual & Visual Representation", "DOK 3 - Error Critique & Proof"])
        btn_gen_assess = st.button("⚡ Build Assessment Slips & Distractor Key", type="primary", use_container_width=True)

    with col_a2:
        if btn_gen_assess:
            if client:
                system_prompt = f"""
                You are a STAAR Math Assessment Specialist for {selected_grade}.
                TASK: Generate a 5-minute Warmup and Exit Slip for {selected_teks}.
                REQUIREMENTS:
                1. Warmup (DOK 2): Error Critique format. Flawed student solution requiring visual CRA model to fix.
                2. Exit Slip (DOK 3): Require algebraic solution AND visual proof + sentence stem justification.
                3. Eduphoria Distractor Flag: Identify Choice A, B, C, D distractor pitfalls.
                """
                response = client.models.generate_content(
                    model="gemini-1.5-pro",
                    contents=f"{system_prompt}\n\nMisconception: {misconception_input}\nDOK: {dok_level}"
                )
                st.markdown(response.text)
            else:
                st.subheader(f"Generated Slips: {selected_teks}")
                st.markdown(f"""
                **WARMUP (Error Critique - DOK 2)**
                > *Student A claims that solving $-2(x - 5) = 14$ gives $x = -2$.*  
                > **Task:** Draw an area model for $-2(x - 5)$ and explain why Student A's constant term is incorrect.
                
                **EXIT SLIP (Proof & Justification - DOK 3)**
                > Solve $-3x + 4 > 16$. Plot your solution on a number line.  
                > **Sentence Stem:** *"I know my inequality sign must point (left/right) because testing $x=0$ results in..."*
                
                **⚠️ Eduphoria Distractor Analysis Key:**
                * **Correct Answer:** $x < -4$
                * **Distractor B ($x > -4$):** Student solved algebra correctly but failed to reverse sign during negative division.
                * **Distractor C ($x < -6$):** Arithmetic sign error when subtracting 4 from 16.
                """)

# ------------------------------------------------------------------------------
# TAB 3: EDUPHORIA DATA & STAAR COACH (FERPA PROTECTED)
# ------------------------------------------------------------------------------
with tab_data:
    st.subheader("Eduphoria Aware CSV Ingestion & STAAR Accountability")
    st.info("🛡️ **FERPA Privacy Shield Active:** Uploaded files are automatically scrubbed of Student Names, Local IDs, and Teacher Names in memory before processing.")

    uploaded_csv = st.file_uploader("Upload Eduphoria Aware Export (.csv):", type=["csv"])
    col_d1, col_d2 = st.columns([1, 1])

    with col_d1:
        st.markdown("##### Processed & Anonymized Benchmark Data")
        if uploaded_csv:
            try:
                raw_df = pd.read_csv(uploaded_csv)
                sanitized_df, dropped_pii_cols = sanitize_eduphoria_dataframe(raw_df)
                
                if dropped_pii_cols:
                    st.success(f"🔒 FERPA Scrub Complete! Removed PII columns: `{', '.join(dropped_pii_cols[:4])}...`")
                
                st.dataframe(sanitized_df.head(5), use_container_width=True)
                
                score_col = next((c for c in sanitized_df.columns if 'score' in c.lower() or 'percent' in c.lower()), None)
                if score_col:
                    dnm_pct = (sanitized_df[score_col] < 60).mean() * 100
                    app_pct = ((sanitized_df[score_col] >= 60) & (sanitized_df[score_col] < 75)).mean() * 100
                    meets_pct = ((sanitized_df[score_col] >= 75) & (sanitized_df[score_col] < 90)).mean() * 100
                    masters_pct = (sanitized_df[score_col] >= 90).mean() * 100
                else:
                    dnm_pct, app_pct, meets_pct, masters_pct = 25.0, 40.0, 25.0, 10.0
            except Exception as e:
                st.error(f"Error parsing file: {e}")
                dnm_pct, app_pct, meets_pct, masters_pct = 25.0, 40.0, 25.0, 10.0
        else:
            dnm_pct = st.number_input("% Did Not Meet", value=25.0)
            app_pct = st.number_input("% Approaches Grade Level", value=40.0)
            meets_pct = st.number_input("% Meets Grade Level", value=25.0)
            masters_pct = st.number_input("% Masters Grade Level", value=10.0)

        approaches_or_above = app_pct + meets_pct + masters_pct
        domain_1_score = (approaches_or_above + meets_pct + masters_pct) / 3
        st.metric(label="Estimated Domain 1 STAAR Component Score", value=f"{domain_1_score:.1f}%")
        
        lowest_teks_summary = st.text_area(
            "Lowest Performing Question / TEKS Summary:",
            value=f"Question #8 ({selected_teks}): 58% error rate. Students chose Distractor C."
        )
        btn_run_data = st.button("📊 Generate 15-Min Reteach Plan", type="primary", use_container_width=True)

    with col_d2:
        if btn_run_data:
            st.markdown("### ⏱️ 15-Minute Tier-1 Reteach Protocol")
            st.markdown(f"""
            **Focus TEKS:** {selected_teks} | **Domain 1 Score:** {domain_1_score:.1f}%
            
            1. **Minutes 0–3 (Error Critique Warmup):**  
               Display Question #8 distractor mistake. Prompt 4-person Kagan teams:  
               *Sentence Stem:* *"Choice C is incorrect because substituting 0 back into the original inequality gives..."*
            
            2. **Minutes 3–10 (Visual CRA Modeling):**  
               Students use Balance Mats to test values left/right of the boundary point to visually prove why negative scale factors reverse inequalities.
            
            3. **Minutes 10–15 (Exit Check):**  
               Administer a 1-question recheck requiring students to plot their answer and write one sentence justifying their direction.
            """)

# ------------------------------------------------------------------------------
# TAB 4: PLC CLUSTER & GROWTH COACH
# ------------------------------------------------------------------------------
with tab_cluster:
    st.subheader("PLC Cluster Meeting & Teacher Growth Plan Coach")
    st.caption("Audits teacher growth goals and builds supportive 1-on-1 coaching conference debriefs.")

    col_c1, col_c2 = st.columns([1, 1])
    with col_c1:
        goal_input = st.text_area(
            "Paste Teacher Growth Goal or Action Plan:",
            height=150,
            value=f"Goal: I will increase student participation in {selected_grade} by asking more higher-order questions during guided practice."
        )
        btn_cluster = st.button("📋 Audit Goal & Generate 5-Min Agenda", type="primary", use_container_width=True)

    with col_c2:
        if btn_cluster:
            st.markdown("""
            <div class="glow-card">
                <h4>🌟 Affirming Goal Audit</h4>
                <p>This goal demonstrates a strong commitment to elevating student discourse! Focusing on higher-order questioning directly targets NIET Standards for <b>Questioning & Problem Solving</b>.</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("""
            ### 📋 1-on-1 Coaching Conference Agenda (5 Minutes)
            1. **Affirm & Celebrate (1 min):** "I love how intentionally you are working to give students more ownership of the math..."
            2. **SMART Refinement (2 mins):** Strengthen metric: *"Increase DOK 2/3 student-led response time from 3s to 5s in 4 out of 5 math lessons."*
            3. **Action Step (2 mins):** Commit to embedding Kagan 4-person team sentence stems in tomorrow's warmup.
            """)
