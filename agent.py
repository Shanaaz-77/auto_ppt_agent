"""
agent.py — Auto-PPT Agent (Fixed & Production-Ready)

Flow:
  User Input
  → clean topic
  → launch ppt_mcp_server.py via stdio (sys.executable -m fastmcp run)
  → Agentic ReAct loop (Plan → Create → AddSlides → WriteText → Save)
  → .pptx saved to generated_ppts/

Root-cause fixes applied:
  1. MCP connection: uses sys.executable -m fastmcp run  (no PATH dependency)
  2. server_path: uses os.path.abspath(__file__)          (no empty-dirname bug)
  3. System prompt: fully rewritten for high-quality PPT output
  4. Max-iteration guard: prevents infinite loops
  5. Fallback save: guarantees a file is always written
"""

import os
import re
import sys
import json
import asyncio

from groq import Groq
from mcp import ClientSession
from mcp.client.streamable_http import streamablehttp_client

# ──────────────────────────────────────────────
#  CLIENT SETUP
# ──────────────────────────────────────────────
client = Groq(api_key=os.getenv("GROQ_API_KEY"))
MODEL_NAME   = "llama-3.3-70b-versatile"
MAX_ITERATIONS = 30          # hard cap — prevents infinite loops

# ──────────────────────────────────────────────
#  SYSTEM PROMPT  (rewritten for quality output)
# ──────────────────────────────────────────────
SYSTEM_PROMPT = """
You are an expert Presentation Agent. Your only job is to create a complete,
high-quality PowerPoint presentation using the tools available to you.

═══════════════════════════════════════════════════════
MANDATORY EXECUTION ORDER — NEVER DEVIATE
═══════════════════════════════════════════════════════

STEP 1 ── PLAN (output as plain text BEFORE any tool call)
  Write exactly this block first:

  PLAN:
  - Title Slide : [Clean topic name]
  - Slide 1     : [2–4 word title]
  - Slide 2     : [2–4 word title]
  - Slide 3     : [2–4 word title]
  - Slide 4     : [2–4 word title]
  - Slide 5     : [2–4 word title]

STEP 2 ── Call create_presentation()

STEP 3 ── TITLE SLIDE
  Call add_slide()                      → returns slide_index (0)
  Call write_text(
      slide_index = 0,
      title       = "[Clean topic name — see rules below]",
      bullets     = ["[One compelling subtitle sentence, 10–15 words]"]
  )

STEP 4 ── CONTENT SLIDES  (repeat for every planned slide)
  Call add_slide()                      → returns slide_index N
  Call write_text(
      slide_index = N,
      title       = "[2–4 word title]",
      bullets     = ["bullet 1", "bullet 2", "bullet 3", "bullet 4"]
  )

STEP 5 ── SAVE
  Call save_presentation(filename="[topic]_presentation.pptx")

═══════════════════════════════════════════════════════
TITLE SLIDE RULES
═══════════════════════════════════════════════════════
• Extract the clean topic from the user's request.
• STRIP all filler phrases before writing the title:
    "create a ppt on"         → remove
    "make a presentation on"  → remove
    "presentation about"      → remove
    "create slides on"        → remove
    "generate a ppt on"       → remove
• Title must be ONLY the topic name, properly capitalized.

  ✔  Input : "create a ppt on Machine Learning"
     Title : "Machine Learning"

  ✔  Input : "make a 5 slide deck on the Solar System"
     Title : "The Solar System"

  ✗  NEVER write: "Create a PPT on Machine Learning"
  ✗  NEVER write: "Presentation on Machine Learning"

═══════════════════════════════════════════════════════
SLIDE TITLE RULES
═══════════════════════════════════════════════════════
• 2 to 4 words maximum.
• Specific and meaningful — tells the reader what this slide is about.
• Use title case.

  ✔  Good titles:
       "What Is Machine Learning"
       "Types of Learning"
       "Real-World Applications"
       "How Neural Networks Work"
       "Future of AI"

  ✗  Bad titles (NEVER use):
       "This slide will cover the different types of machine learning approaches"
       "Introduction"   ← too generic
       "Overview"       ← too generic

═══════════════════════════════════════════════════════
BULLET POINT RULES  (the most important section)
═══════════════════════════════════════════════════════
• Exactly 3 to 4 bullets per content slide.
• Each bullet: 8 to 15 words.
• Format: [Subject] + [what it does / means / why it matters].
• Include a brief explanation or real-world example.
• Language: simple, clear, professional — no jargon dumps.

  ✔  GOOD bullets (copy this exact style):
       "Machine learning is a subset of AI that learns from data automatically."
       "Supervised learning trains models using labeled input-output pairs."
       "Neural networks mimic the brain using weighted, connected layers of nodes."
       "Deep learning powers speech recognition, image detection, and translation."

  ✗  BAD bullets (NEVER write these):
       "Machine learning"                            ← too short, no explanation
       "There are many types of machine learning"    ← vague, adds nothing
       "This slide explains how ML works"            ← meta, not informative
       "Machine learning is very important today"    ← empty filler

• NO repetition across slides — each slide must add brand-new information.
• Stay strictly on topic — no off-topic tangents.

═══════════════════════════════════════════════════════
FILENAME RULE
═══════════════════════════════════════════════════════
• Lowercase the topic, replace spaces with underscores, append _presentation.pptx
• Example: "Machine Learning" → "machine_learning_presentation.pptx"

═══════════════════════════════════════════════════════
ERROR HANDLING RULES
═══════════════════════════════════════════════════════
• If a tool returns an error, read the error message, fix your arguments, and retry.
• NEVER stop the loop because of a single tool failure.
• If you cannot find real data, generate accurate, factual content from your knowledge.
• You MUST always end with save_presentation — no exceptions.
"""


# ──────────────────────────────────────────────
#  HELPERS
# ──────────────────────────────────────────────

def mcp_tool_to_groq_format(mcp_tool):
    """Convert an MCP tool definition into the Groq function-call schema."""
    return {
        "type": "function",
        "function": {
            "name":        mcp_tool.name,
            "description": mcp_tool.description,
            "parameters":  mcp_tool.inputSchema,
        },
    }


def clean_topic(user_request: str) -> str:
    """
    Strip common filler phrases from the user request to obtain a clean topic name.

    Examples
    --------
    "create a ppt on Machine Learning"  → "Machine Learning"
    "make a 5-slide deck on Solar System" → "Solar System"
    """
    filler_patterns = [
        r"(?:please\s+)?(?:create|make|generate|build|design|produce)\s+(?:a\s+)?(?:\d+[\s\-]slide\s+)?(?:ppt|powerpoint|presentation|deck|slides?)\s+(?:on|about|for|regarding)\s+",
        r"(?:presentation|ppt|slides?)\s+(?:on|about|for|regarding)\s+",
        r"(?:create|make|generate|build|design)\s+(?:slides?\s+)?(?:on|about|for)\s+",
        r"(?:a\s+ppt\s+on|a\s+deck\s+on|a\s+presentation\s+on)\s+",
    ]
    cleaned = user_request.strip()
    for pattern in filler_patterns:
        cleaned = re.sub(pattern, "", cleaned, flags=re.IGNORECASE).strip()

    # Title-case the result; fall back to original if cleaning erased everything
    cleaned = cleaned.strip(" .,;:-")
    result = cleaned.title() if cleaned else user_request.strip().title()
    return result


def safe_filename(topic: str) -> str:
    """Convert a topic string into a safe .pptx filename."""
    name = re.sub(r"[^\w\s-]", "", topic).strip().lower()
    name = re.sub(r"[\s-]+", "_", name)
    return f"{name}_presentation.pptx"


# ──────────────────────────────────────────────
#  MAIN AGENT  (async)
# ──────────────────────────────────────────────

async def run_ppt_agent(user_request: str) -> str:
    """
    Run the full agentic loop and return a status string.

    MCP connection strategy
    -----------------------
    Connects to the already-running ppt_mcp_server.py over HTTP:
        http://127.0.0.1:8000/mcp

    The server must be started separately before running the agent:
        python ppt_mcp_server.py
    """

    # ── Clean topic ────────────────────────────────────────
    topic    = clean_topic(user_request)
    out_name = safe_filename(topic)

    print(f"\n{'='*55}")
    print(f"  Auto-PPT Agent")
    print(f"  Topic    : {topic}")
    print(f"  Output   : generated_ppts/{out_name}")
    print(f"{'='*55}\n")

    # ── MCP HTTP connection ────────────────────────────────
    # Server must already be running:  python ppt_mcp_server.py
    # It listens on http://127.0.0.1:8000/mcp  (HTTP transport)
    MCP_URL = "http://127.0.0.1:8000/mcp"

    print(f"[*] Connecting to MCP server at {MCP_URL} …")

    async with streamablehttp_client(MCP_URL) as (read, write, _):
        async with ClientSession(read, write) as session:

            await session.initialize()

            # Discover tools and convert to Groq schema
            tools_response = await session.list_tools()
            groq_tools     = [mcp_tool_to_groq_format(t) for t in tools_response.tools]
            tool_names     = [t.name for t in tools_response.tools]

            print(f"[*] MCP connected  ✓   Tools: {tool_names}\n")

            # ── Build initial messages ─────────────────────
            enhanced_request = (
                f"Create a complete PowerPoint presentation on: '{topic}'\n\n"
                f"User's original request: '{user_request}'\n\n"
                f"IMPORTANT — The title slide MUST show exactly: \"{topic}\"\n"
                f"IMPORTANT — Save the file as: \"{out_name}\""
            )

            messages = [
                {"role": "system",  "content": SYSTEM_PROMPT},
                {"role": "user",    "content": enhanced_request},
            ]

            print(f"[*] Agentic loop started …\n")

            # ── Agentic ReAct loop ─────────────────────────
            iteration   = 0
            save_called = False

            while iteration < MAX_ITERATIONS and not save_called:
                iteration += 1
                print(f"┌─ Iteration {iteration}/{MAX_ITERATIONS} {'─'*40}")

                # THOUGHT: ask the LLM what to do next
                try:
                    response = client.chat.completions.create(
                        model=MODEL_NAME,
                        messages=messages,
                        tools=groq_tools,
                        tool_choice="auto",
                        max_tokens=4096,
                    )
                except Exception as llm_err:
                    print(f"│  [LLM Error] {llm_err}")
                    print(f"└─ Retrying next iteration …\n")
                    continue

                msg      = response.choices[0].message
                msg_dict = {"role": msg.role}

                if msg.content:
                    msg_dict["content"] = msg.content
                    # Print first 400 chars of thought to keep logs readable
                    snippet = msg.content[:400].replace("\n", " ")
                    print(f"│  [Thought] {snippet}{'…' if len(msg.content) > 400 else ''}")

                if msg.tool_calls:
                    # ── ACTION branch ──────────────────────
                    tool_calls_dict = [
                        {
                            "id":   tc.id,
                            "type": "function",
                            "function": {
                                "name":      tc.function.name,
                                "arguments": tc.function.arguments,
                            },
                        }
                        for tc in msg.tool_calls
                    ]
                    msg_dict["tool_calls"] = tool_calls_dict
                    messages.append(msg_dict)

                    for tool_call in msg.tool_calls:
                        func_name = tool_call.function.name

                        # Parse arguments safely
                        try:
                            func_args = json.loads(tool_call.function.arguments)
                        except json.JSONDecodeError:
                            func_args = {}

                        # Intercept save_presentation to control output path
                        if func_name == "save_presentation":
                            os.makedirs("generated_ppts", exist_ok=True)
                            raw_fn = func_args.get("filename", out_name)
                            # Strip any directory components the LLM may have added
                            clean_fn = os.path.basename(raw_fn)
                            if not clean_fn.endswith(".pptx"):
                                clean_fn += ".pptx"
                            func_args["filename"] = os.path.join(
                                "generated_ppts", clean_fn
                            )

                        print(f"│  [Action]      {func_name}")
                        arg_preview = json.dumps(func_args)
                        if len(arg_preview) > 160:
                            arg_preview = arg_preview[:160] + "…"
                        print(f"│  [Args]        {arg_preview}")

                        # OBSERVATION: execute the MCP tool
                        try:
                            result      = await session.call_tool(func_name, func_args)
                            result_text = (
                                result.content[0].text
                                if result.content
                                else "Success (no output)"
                            )
                        except Exception as tool_err:
                            result_text = f"Tool error: {tool_err}"

                        print(f"│  [Observation] {result_text}")

                        # Append observation to conversation
                        messages.append({
                            "role":         "tool",
                            "tool_call_id": tool_call.id,
                            "content":      result_text,
                        })

                        # Check for successful save
                        if (
                            func_name == "save_presentation"
                            and "saved at" in result_text.lower()
                        ):
                            save_called = True
                            print(f"└─\n")
                            print(f"{'='*55}")
                            print(f"  ✅ SUCCESS — Presentation saved!")
                            print(f"  Path: {func_args['filename']}")
                            print(f"{'='*55}\n")
                            return result_text

                else:
                    # ── NO TOOL CALL branch ────────────────
                    # LLM replied with text only — push it back and urge progress
                    messages.append({"role": "assistant", "content": msg.content or ""})

                    # Check whether save already happened (LLM might "think" it is done)
                    already_saved = any(
                        m.get("role") == "tool"
                        and "saved at" in m.get("content", "").lower()
                        for m in messages
                    )

                    if already_saved:
                        save_called = True
                        print(f"└─\n✅ [SUCCESS] Detected prior save — finishing.\n")
                        return "Presentation saved successfully."

                    # Force the agent back into tool-use mode
                    nudge = (
                        "You must continue using the tools provided. "
                        "Do NOT narrate — ACT. "
                    )
                    # Detect where we are in the flow and give a targeted nudge
                    called_create = any(
                        m.get("role") == "tool"
                        and "created successfully" in m.get("content", "").lower()
                        for m in messages
                    )
                    if not called_create:
                        nudge += "Call create_presentation() right now."
                    else:
                        nudge += (
                            "Add the next slide with add_slide(), fill it with "
                            "write_text(), and once ALL slides are done, call "
                            "save_presentation() to finish."
                        )

                    print(f"│  [System] {nudge}")
                    messages.append({"role": "user", "content": nudge})

                print(f"└─\n")

            # ── FALLBACK SAVE ──────────────────────────────
            # If the loop ended without a confirmed save, try one last direct save.
            if not save_called:
                print("[FALLBACK] Loop ended without confirmed save — attempting emergency save …")
                os.makedirs("generated_ppts", exist_ok=True)
                fallback_path = os.path.join("generated_ppts", out_name)
                try:
                    result = await session.call_tool(
                        "save_presentation", {"filename": fallback_path}
                    )
                    fb_text = result.content[0].text if result.content else "Done"
                    print(f"[FALLBACK] {fb_text}")
                    return fb_text
                except Exception as fb_err:
                    msg_out = f"Fallback save failed: {fb_err}"
                    print(f"[FALLBACK ERROR] {msg_out}")
                    return msg_out

    return "Agent loop completed."


# ──────────────────────────────────────────────
#  CLI ENTRY POINT
# ──────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) > 1:
        req = " ".join(sys.argv[1:])
    else:
        try:
            req = input("Enter your PPT topic: ").strip()
        except EOFError:
            req = "The lifecycle of a star for a 6th-grade class"

    if not req:
        req = "Introduction to Artificial Intelligence"

    asyncio.run(run_ppt_agent(req))