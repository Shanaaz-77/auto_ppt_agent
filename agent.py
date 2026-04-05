import os
import sys
import json
import asyncio
from groq import Groq
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

# Initialize Groq client
# We expect GROQ_API_KEY to be set in the environment
client = Groq(api_key=os.getenv("GROQ_API_KEY"))
MODEL_NAME = "llama-3.3-70b-versatile" 

SYSTEM_PROMPT = """
You are a Presentation Agent. Your job is to create a complete PowerPoint presentation based on the user's topic.
You have access to tools to interact with a PowerPoint MCP server.

CRITICAL INSTRUCTIONS - YOU MUST FOLLOW THESE STEPS:
Step 1: Plan slide titles as an array (Think step: Decide the outline).
Step 2: Initialize the presentation by calling 'create_presentation'.
Step 3: For each planned title, in order:
    a. Call 'add_slide' to get a new slide index.
    b. Output 3-5 bullet points and the title, then call 'write_text' with the slide_index.
    c. (Optional) Provide a URL and use 'add_image' to make it look nice. If you fail to find one, skip or hallucinate gracefully.
Step 4: Once all slides are complete, call 'save_presentation' with a distinct filename (e.g. 'topic_slides.pptx').

RULES:
- Never skip the planning step. 
- If a tool fails, gracefully handle it, adjust plan or hallucinate plausible content, and keep going without crashing.
- Do not make assumptions, use tools to advance the presentation state.
"""

def mcp_tool_to_groq_format(mcp_tool):
    """Converts an MCP tool definition to Groq/OpenAI function schema."""
    return {
        "type": "function",
        "function": {
            "name": mcp_tool.name,
            "description": mcp_tool.description,
            "parameters": mcp_tool.inputSchema
        }
    }

async def run_ppt_agent(user_request: str):
    server_path = os.path.join(os.path.dirname(__file__), "ppt_mcp_server.py")
    
    fastmcp_cmd = os.path.join(os.path.dirname(sys.executable), "fastmcp")
    if os.name == 'nt':
        fastmcp_cmd += ".exe"
    if not os.path.exists(fastmcp_cmd):
        fastmcp_cmd = "fastmcp"
        
    # We run the FastMCP app directly over stdio
    server_params = StdioServerParameters(
        command=fastmcp_cmd,
        args=["run", f"{server_path}:mcp"]
    )
    
    print(f"[*] Starting MCP server from {server_path}...")
    
    async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            
            # Fetch tools safely
            mcp_tools_response = await session.list_tools()
            groq_tools = [mcp_tool_to_groq_format(t) for t in mcp_tools_response.tools]
            print(f"[*] Connected to MCP Server. Available tools: {[t.name for t in mcp_tools_response.tools]}")
            
            messages = [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": f"Please create a presentation on: {user_request}"}
            ]
            
            print(f"[*] Starting Agentic Loop for topic: {user_request}\n")
            
            while True:
                # 1. Thought / Action Decision
                response = client.chat.completions.create(
                    model=MODEL_NAME,
                    messages=messages,
                    tools=groq_tools,
                    tool_choice="auto",
                    max_tokens=4096
                )
                
                msg = response.choices[0].message
                
                # We cannot append the raw object directly, we must convert it to a dict properly for the Groq client
                # Especially handling None values
                msg_dict = {"role": msg.role}
                if msg.content:
                    msg_dict["content"] = msg.content
                    print(f"  [Thought] {msg.content}")
                
                if msg.tool_calls:
                    # Convert tool calls to dict
                    tool_calls_dict = []
                    for t in msg.tool_calls:
                        tool_calls_dict.append({
                            "id": t.id,
                            "type": "function",
                            "function": {
                                "name": t.function.name,
                                "arguments": t.function.arguments
                            }
                        })
                    msg_dict["tool_calls"] = tool_calls_dict
                    messages.append(msg_dict)
                    
                    for tool_call in msg.tool_calls:
                        func_name = tool_call.function.name
                        try:
                            func_args = json.loads(tool_call.function.arguments)
                        except json.JSONDecodeError:
                            func_args = {}
                            
                        
                        if func_name == "save_presentation":
                            os.makedirs("generated_ppts", exist_ok=True)
                            filename = func_args.get("filename", "output.pptx")
                            filename = os.path.basename(filename) # ensure no other directories
                            func_args["filename"] = os.path.join("generated_ppts", filename)
                            
                        print(f"  [Action] Calling '{func_name}' with args {func_args}")
                        
                        try:
                            # 2. Observation
                            result = await session.call_tool(func_name, func_args)
                            result_text = result.content[0].text if result.content else "Success"
                        except Exception as e:
                            # Graceful fallback on crash
                            result_text = f"Error: {str(e)}"
                            
                        print(f"  [Observation] {result_text}")
                            
                        # Append observation mapping to the tool_call id
                        messages.append({
                            "role": "tool",
                            "tool_call_id": tool_call.id,
                            "content": result_text
                        })
                        
                        # Completion check
                        if func_name == "save_presentation" and not "Error" in result_text:
                            print(f"\n[SUCCESS] Presentation finished successfully!")
                            return
                else:
                    # No tool calls made.
                    messages.append({"role": "assistant", "content": msg.content or ""})
                    
                    # Force continuation if save_presentation hasn't been called
                    saved = any(
                        m.get("role") == "tool" and "Presentation saved at" in m.get("content", "")
                        for m in messages
                    )
                    
                    if not saved:
                        print("  [System] Forcing agent to continue until save_presentation is called...")
                        messages.append({
                            "role": "user", 
                            "content": "Please continue and use the available tools to complete the presentation. Call 'save_presentation' when you are totally done."
                        })
                    else:
                        print(f"\n[SUCCESS] Presentation finished successfully!")
                        return

if __name__ == "__main__":
    if len(sys.argv) > 1:
        req = " ".join(sys.argv[1:])
    else:
        try:
            req = input("Enter a topic for the PPT: ")
        except EOFError:
            req = "The lifecycle of a star suitable for a 6th grade class"
        
    asyncio.run(run_ppt_agent(req))