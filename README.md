# 🎯 Auto PPT Generator  
### An Agentic AI System Powered by MCP (Model Context Protocol)

---

## 🚀 Overview

**Auto PPT Generator** is a fully automated AI agent that transforms a single-line user prompt into a complete, structured PowerPoint presentation.

It demonstrates a real-world implementation of:
- Agentic AI (ReAct-style reasoning)
- MCP (Model Context Protocol)
- Tool-driven execution
- LLM-powered content generation

---

## 🧠 What Makes This Project Special?

Unlike traditional scripts, this system behaves like an **intelligent agent** that:

1. Understands user intent  
2. Plans presentation structure  
3. Dynamically calls tools  
4. Generates content  
5. Produces a final `.pptx` file  

All without manual intervention.

---

## ⚙️ Core Architecture

    🧑 User Input (Streamlit UI)
                 ↓
        🎯 Agent (agent.py)
    (Planning + Reasoning Layer)
                 ↓
     🔌 MCP Client (inside agent.py)
                 ↓
     🧰 MCP Server (ppt_mcp_server.py)
                 ↓
    📊 PowerPoint File (.pptx)

---

## 🔌 MCP Integration (Key Highlight)

This project strictly follows the **MCP architecture**, separating reasoning from execution.

### 🟢 MCP Server (`ppt_mcp_server.py`)
Provides modular tools:

- `create_presentation` → Initialize PPT  
- `add_slide` → Add slide layout  
- `write_text` → Insert title + bullets  
- `add_image` → Add visual content  
- `save_presentation` → Save file  

👉 The server acts as a **tool provider**.

---

### 🔵 MCP Client (`agent.py`)
The agent connects via MCP and:

- Calls tools dynamically  
- Passes structured arguments  
- Receives responses (observations)  

👉 The client acts as a **decision-maker**.

---

## 🤖 Agentic Workflow (ReAct Loop)

The system follows a reasoning loop:


User Prompt
↓
Plan Slides (LLM)
↓
[Loop]
Thought → "Need new slide"
Action → add_slide
Observation → slide index

Thought → "Generate content"
Action → write_text
Observation → content added
[/Loop]
↓
Save Presentation


👉 This ensures **dynamic, intelligent execution** instead of hardcoded logic.

---

## 🧠 LLM Integration

- Powered by **Groq (LLaMA 3.1)**
- Used for:
  - Slide planning
  - Content generation
  - Bullet optimization

---

## 🎨 Frontend (Streamlit)

A clean and minimal UI that allows users to:

- Enter a prompt  
- Generate presentation  
- Download PPT instantly  

---

## 📁 File Structure


auto-ppt-generator/
│
├── agent.py # AI Agent (MCP client + logic)
├── ppt_mcp_server.py # MCP Server (tools)
├── app.py # Streamlit frontend
├── generated_ppts/ # Output PPT files
├── requirements.txt
└── README.md


---

## ⚡ Features

- 🤖 Agent-based architecture (not script-based)  
- 🔌 True MCP integration (client + server separation)  
- 🧠 Intelligent slide planning  
- 📊 Auto-generated structured slides  
- 🖼️ Image integration  
- 🎯 Dynamic slide count (user-defined)  
- 🎨 Clean UI (Streamlit)  
- 📁 Organized file output  

---

## 🛠️ Installation

### 1. Clone Repository

```bash
git clone https://github.com/Shanaaz-77/auto-ppt-generator.git
cd auto-ppt-generator
