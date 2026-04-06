"""
app.py — Auto PPT Generator
Streamlit frontend for the Auto-PPT Agent.

Flow: UI → run_ppt_agent (agent.py) → MCP Client → MCP Server → .pptx
"""

import os
import glob
import asyncio
import streamlit as st
from agent import run_ppt_agent

# ─────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────
st.set_page_config(
    page_title="Auto PPT Generator",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────
#  STYLES
# ─────────────────────────────────────────
st.markdown("""
<style>
    /* ── Hide sidebar & Streamlit chrome ── */
    [data-testid="stSidebar"]          { display: none !important; }
    [data-testid="stDecoration"]       { display: none !important; }
    #MainMenu, footer, header          { visibility: hidden; }

    /* ── Page background ── */
    .stApp {
        background-color: #f6f8fb;
    }

    /* ── Center block ── */
    .block-container {
        max-width: 680px;
        padding-top: 3.5rem;
        padding-bottom: 3rem;
    }

    /* ── Card wrapper ── */
    .card {
        background: #ffffff;
        border-radius: 20px;
        padding: 2.5rem 2rem;
        box-shadow: 0 4px 24px rgba(0, 0, 0, 0.07);
        margin-top: 1.5rem;
    }

    /* ── Header ── */
    .main-title {
        text-align: center;
        font-size: 2.4rem;
        font-weight: 800;
        color: #1a1a2e;
        letter-spacing: -0.5px;
        margin-bottom: 0.3rem;
    }
    .subtitle {
        text-align: center;
        font-size: 1.05rem;
        color: #6b7280;
        margin-bottom: 0;
    }

    /* ── Text input ── */
    .stTextInput > div > div > input {
        background-color: #ffffff !important;
        border-radius: 12px !important;
        border: 1.5px solid #e5e7eb !important;
        padding: 0.75rem 1rem !important;
        font-size: 0.97rem !important;
        color: #1a1a2e !important;
        box-shadow: none !important;
    }
    .stTextInput > div > div > input:focus {
        border-color: #4f46e5 !important;
        box-shadow: 0 0 0 3px rgba(79, 70, 229, 0.12) !important;
    }
    .stTextInput label {
        font-weight: 600;
        color: #374151;
        font-size: 0.95rem;
    }

    /* ── Primary button ── */
    div[data-testid="stButton"] > button {
        width: 100%;
        background: linear-gradient(135deg, #4f46e5, #7c3aed) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.75rem 1.5rem !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.3px !important;
        cursor: pointer !important;
        transition: opacity 0.2s ease, transform 0.1s ease !important;
        box-shadow: 0 4px 12px rgba(79, 70, 229, 0.3) !important;
    }
    div[data-testid="stButton"] > button:hover {
        opacity: 0.92 !important;
        transform: translateY(-1px) !important;
    }
    div[data-testid="stButton"] > button:active {
        transform: translateY(0px) !important;
    }

    /* ── Download button ── */
    div[data-testid="stDownloadButton"] > button {
        width: 100%;
        background: #ffffff !important;
        color: #4f46e5 !important;
        border: 2px solid #4f46e5 !important;
        border-radius: 12px !important;
        padding: 0.65rem 1.5rem !important;
        font-size: 0.97rem !important;
        font-weight: 600 !important;
        cursor: pointer !important;
        transition: background 0.2s ease !important;
    }
    div[data-testid="stDownloadButton"] > button:hover {
        background: #f0f0ff !important;
    }

    /* ── Success / error alerts ── */
    .stAlert {
        border-radius: 12px !important;
    }

    /* ── Spinner text ── */
    .stSpinner > div {
        font-size: 0.95rem;
        color: #6b7280;
    }

    /* ── Divider ── */
    .divider {
        border: none;
        border-top: 1px solid #f0f0f5;
        margin: 1.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
#  HEADER
# ─────────────────────────────────────────
st.markdown('<p class="main-title">⚡ Auto PPT Generator</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Create presentations instantly — powered by AI & MCP</p>', unsafe_allow_html=True)

st.markdown("<div style='height: 1.5rem'></div>", unsafe_allow_html=True)

# ─────────────────────────────────────────
#  GROQ API KEY CHECK
# ─────────────────────────────────────────
if not os.getenv("GROQ_API_KEY"):
    st.error(
        "⚠️ **GROQ_API_KEY is not set.**\n\n"
        "Please set it before running:\n"
        "```\nset GROQ_API_KEY=your_key_here\n```"
    )
    st.stop()

# ─────────────────────────────────────────
#  MAIN CARD
# ─────────────────────────────────────────
with st.container():
    # Input
    prompt = st.text_input(
        "Enter your topic",
        placeholder="e.g, Create a 5 slide presentation on Artificial Intelligence",
    )

    st.markdown("<div style='height: 0.6rem'></div>", unsafe_allow_html=True)

    # Generate button
    generate_clicked = st.button("⚡ Generate PPT", use_container_width=True)

# ─────────────────────────────────────────
#  GENERATION LOGIC
# ─────────────────────────────────────────
if generate_clicked:
    # Validate input
    if not prompt.strip():
        st.error("⚠️ Please enter a topic before generating.")
        st.stop()

    # Snapshot existing files before generation
    os.makedirs("generated_ppts", exist_ok=True)
    before_files = set(glob.glob("generated_ppts/*.pptx"))

    try:
        with st.spinner("Generating your PPT..."):
            asyncio.run(run_ppt_agent(prompt.strip()))

        # Detect newly created file
        after_files = set(glob.glob("generated_ppts/*.pptx"))
        new_files = after_files - before_files

        if new_files:
            ppt_path = sorted(new_files, key=os.path.getmtime, reverse=True)[0]
        else:
            # Fallback: pick the most recently modified file
            all_files = glob.glob("generated_ppts/*.pptx")
            ppt_path = max(all_files, key=os.path.getmtime) if all_files else None

        st.markdown("<div style='height: 0.8rem'></div>", unsafe_allow_html=True)
        st.success("✅ Your PPT has been generated successfully!")

        if ppt_path and os.path.exists(ppt_path):
            with open(ppt_path, "rb") as f:
                ppt_bytes = f.read()

            st.markdown("<div style='height: 0.5rem'></div>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Download PPT",
                data=ppt_bytes,
                file_name=os.path.basename(ppt_path),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        else:
            st.warning("⚠️ Presentation was generated but the file could not be located.")

    except Exception as e:
        st.error(f"❌ Something went wrong: {str(e)}")
