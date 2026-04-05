import os
import asyncio
import glob
import threading

import streamlit as st
from agent import run_ppt_agent

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Auto PPT Generator",
    page_icon="📊",
    layout="centered",
)

# ── Minimal, safe CSS (no input colour overrides that break visibility) ───────
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    /* Soft dark page background */
    .stApp { background: #1a1a2e; }

    /* Card wrapper */
    .card {
        background: rgba(255,255,255,0.06);
        border: 1px solid rgba(255,255,255,0.12);
        border-radius: 18px;
        padding: 2.5rem 2rem 2rem;
        margin-top: 2rem;
    }

    /* Gradient title */
    .title {
        font-size: 2.3rem;
        font-weight: 700;
        text-align: center;
        background: linear-gradient(90deg,#a78bfa,#60a5fa,#34d399);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: .25rem;
    }
    .subtitle {
        text-align: center;
        color: rgba(255,255,255,.45);
        font-size: .92rem;
        margin-bottom: 1.8rem;
    }

    /* Label colour */
    label { color: rgba(255,255,255,.75) !important; }

    /* Generate button */
    div.stButton > button {
        background: linear-gradient(90deg,#7c3aed,#4f46e5) !important;
        color: #fff !important;
        border: none !important;
        border-radius: 12px !important;
        padding: .7rem 1.5rem !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        width: 100% !important;
        transition: opacity .2s, transform .15s !important;
    }
    div.stButton > button:hover {
        opacity: .85 !important;
        transform: translateY(-1px) !important;
    }

    /* Download button */
    div.stDownloadButton > button {
        background: linear-gradient(90deg,#059669,#0d9488) !important;
        color: #fff !important;
        border: none !important;
        border-radius: 12px !important;
        padding: .7rem 1.5rem !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        width: 100% !important;
    }

    /* Hide Streamlit chrome */
    #MainMenu, footer, header { visibility: hidden; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Session state ─────────────────────────────────────────────────────────────
if "pptx_path" not in st.session_state:
    st.session_state.pptx_path = None
if "error_msg" not in st.session_state:
    st.session_state.error_msg = None

# ── Helper ────────────────────────────────────────────────────────────────────

def _run_agent_in_thread(prompt: str, result: dict) -> None:
    """
    Target for a dedicated thread.
    Creates a BRAND-NEW event loop — completely isolated from Streamlit's
    tornado/asyncio infrastructure — then runs the agent to completion.
    Any exception is stored in result['error'] for the main thread to raise.
    """
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        loop.run_until_complete(run_ppt_agent(prompt))
    except Exception as exc:
        result["error"] = exc
    finally:
        loop.close()


def generate_presentation(prompt: str) -> str:
    """Kick off the agent in a dedicated thread and return the new .pptx path."""
    before = set(glob.glob(os.path.join("generated_ppts", "*.pptx")))

    result: dict = {"error": None}
    thread = threading.Thread(target=_run_agent_in_thread, args=(prompt, result), daemon=True)
    thread.start()
    thread.join()  # block Streamlit until agent is done (spinner runs via st.spinner)

    # Re-raise any exception from the agent thread
    if result["error"] is not None:
        raise result["error"]

    after = set(glob.glob(os.path.join("generated_ppts", "*.pptx")))
    new = after - before
    if new:
        return new.pop()

    # Fallback: most-recently modified file
    files = glob.glob(os.path.join("generated_ppts", "*.pptx"))
    return max(files, key=os.path.getmtime) if files else None

# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown('<div class="card">', unsafe_allow_html=True)

st.markdown('<p class="title">📊 Auto PPT Generator</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="subtitle">Describe your presentation — the AI agent builds it for you.</p>',
    unsafe_allow_html=True,
)

# Prompt input — plain Streamlit widget, NO colour override so text is always visible
prompt = st.text_area(
    "Enter your prompt",
    placeholder='e.g. "Create a 5-slide presentation on Artificial Intelligence"',
    height=120,
)

if st.button("✨ Generate PPT", use_container_width=True):
    # Reset previous result
    st.session_state.pptx_path = None
    st.session_state.error_msg = None

    if not prompt.strip():
        st.session_state.error_msg = "⚠️ Please enter a prompt before generating."
    else:
        with st.spinner("🤖 Agent is working… this may take a minute or two."):
            try:
                path = generate_presentation(prompt.strip())
                if path and os.path.exists(path):
                    st.session_state.pptx_path = path
                else:
                    st.session_state.error_msg = (
                        "⚠️ Agent finished but no .pptx file was found in generated_ppts/."
                    )
            except Exception as exc:
                st.session_state.error_msg = f"❌ Something went wrong:\n\n{exc}"

# ── Results (always rendered, driven by session state) ────────────────────────

if st.session_state.error_msg:
    st.error(st.session_state.error_msg)

if st.session_state.pptx_path:
    fname = os.path.basename(st.session_state.pptx_path)
    st.success(f"✅ Presentation ready: **{fname}**")

    with open(st.session_state.pptx_path, "rb") as f:
        st.download_button(
            label="⬇️ Download Presentation",
            data=f.read(),          # read all bytes before widget renders
            file_name=fname,
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".presentationml.presentation"
            ),
            use_container_width=True,
        )

st.markdown("</div>", unsafe_allow_html=True)
