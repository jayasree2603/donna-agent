# donna_agent.py
import re
import os
import time
import asyncio
import logging
import sys
from docx import Document
from dotenv import load_dotenv
from typing import Dict, List, Optional, Tuple

# --- USE THE IMPORTS THAT WORK FOR YOU ---
from google.adk.agents import Agent
from google.adk.sessions import InMemorySessionService
from google.adk.tools import FunctionTool
from google.adk.models.lite_llm import LiteLlm
from google.adk.runners import Runner
from google.genai import types

load_dotenv()

# --- Step 1: Define the Tool Functions ---

def find_template(user_query: str, directory_path: str) -> dict:
    """
    Find the template that matches the user's query.
    
    Args:
        user_query: The user's request describing what template they want
        directory_path: Path to the directory containing templates
        
    Returns:
        Dictionary with template information or error message
    """
    if not os.path.isdir(directory_path):
        return {"error": f"Directory not found: {directory_path}"}
    
    templates = {}
    try:
        for filename in os.listdir(directory_path):
            if filename.endswith("_template.docx"):
                friendly_name = filename.replace("_template.docx", "").replace("_", " ")
                templates[friendly_name.lower()] = os.path.join(directory_path, filename)
    except Exception as e:
        return {"error": f"Could not read directory: {e}"}
    
    # Try to find a direct match
    query_lower = user_query.lower()
    if query_lower in templates:
        return {
            "template_name": query_lower,
            "template_path": templates[query_lower],
            "status": "match_found"
        }
    
    # Try to find a partial match
    for template_name, template_path in templates.items():
        if query_lower in template_name or template_name in query_lower:
            return {
                "template_name": template_name,
                "template_path": template_path,
                "status": "partial_match"
            }
    
    # No match found, return available templates
    return {
        "status": "no_match",
        "available_templates": list(templates.keys()),
        "message": "No matching template found. Please choose from the available templates."
    }

def get_variables(template_path: str) -> dict:
    """
    Extract variables from a template and prompt the user for their values.
    
    Args:
        template_path: Path to the template file
        
    Returns:
        Dictionary with variables and their values
    """
    try:
        doc = Document(template_path)
        full_text = []
        for para in doc.paragraphs: 
            full_text.append(para.text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs: 
                        full_text.append(para.text)
        
        # Updated regex pattern to match single curly braces
        placeholders = re.findall(r'\{(\w+)\}', "\n".join(full_text))
        unique_placeholders = list(set(placeholders))
        
        if not unique_placeholders:
            return {"status": "no_variables", "message": "No variables found in the template."}
        
        return {
            "status": "variables_found",
            "variables": unique_placeholders,
            "message": f"Found {len(unique_placeholders)} variables in the template. Please provide values for each variable."
        }
    except Exception as e: 
        return {"status": "error", "message": f"Error reading the .docx file: {e}"}

def fill_template(template_path: str, variables: dict) -> dict:
    """
    Fill the template with the provided variables and generate the document.
    
    Args:
        template_path: Path to the template file
        variables: Dictionary of variable names and their values
        
    Returns:
        Dictionary with status and path to the generated document
    """
    try:
        doc = Document(template_path)
        
        # Replace in paragraphs
        for para in doc.paragraphs:
            for run in para.runs:
                for key, value in variables.items():
                    placeholder = f"{{{key}}}"
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(value))
        
        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            for key, value in variables.items():
                                placeholder = f"{{{key}}}"
                                if placeholder in run.text:
                                    run.text = run.text.replace(placeholder, str(value))
        
        # Ensure results directory exists
        results_dir = "results"
        os.makedirs(results_dir, exist_ok=True)
        
        # Generate a unique filename
        base_filename = os.path.basename(template_path).replace("_template.docx", "")
        timestamp = int(time.time())
        new_filename = f"{base_filename}_{timestamp}.docx"
        new_filepath = os.path.join(results_dir, new_filename)
        
        # Save the document in the results folder
        doc.save(new_filepath)
        
        return {
            "status": "success",
            "message": f"Document generated successfully. The file is saved as: {new_filepath}",
            "file_path": new_filepath,
            "filename": new_filename
        }
        
    except Exception as e: 
        return {"status": "error", "message": f"Error generating the document: {e}"}
# --- Step 2: Create FunctionTool Objects from the functions ---
find_template_tool = FunctionTool(find_template)
get_variables_tool = FunctionTool(get_variables)
fill_template_tool = FunctionTool(fill_template)

# --- Step 3: Create the Agent with the Correct Model and Tools ---
# Initialize LiteLlm with a different approach
model = LiteLlm(model="gpt-4o-mini")

instruction = """
You are Donna, an intelligent HR Document Assistant. Your primary ability is to discover and use any available document template.
Your precise workflow is:
1.  **Find Template:** Use the find_template tool with the user's query and './templates' as the directory path to find the appropriate template.
2.  **Get Variables:** If a template is found, use the get_variables tool to extract variables from the template.
3.  **Collect Variable Values:** Ask the user for each variable's value.
4.  **Fill Template:** Use the fill_template tool with the template path and the collected variables.
5.  **Present Result:** Relay the success message to the user, including the full path to the generated document in the results folder.

IMPORTANT: Always use './templates' as the directory path when calling find_template.
The templates use variables in the format like company_name and name with single curly braces.
All generated documents are saved in the 'results' folder.
"""
donna_agent = Agent(
    name="Donna",
    model=model,
    tools=[find_template_tool, get_variables_tool, fill_template_tool],
    instruction=instruction
)

# --- Step 4: Setup Session Service and Runner ---
APP_NAME = "Donna-HR-Agent"
USER_ID = "local-user"

session_service = InMemorySessionService()

runner = Runner(
    agent=donna_agent,
    app_name=APP_NAME,
    session_service=session_service
)

# --- Step 5: Asynchronous Interaction Loop ---
async def run_donna_assistant():
    print("Donna, the HR Document Assistant. Type 'quit' to exit.")
    
    initial_message = "Hello, I'm Donna! What document would you like to create today?"
    print(f"Agent: {initial_message}")
    
    # Create a session using 'await' since we are in an async function
    session = await session_service.create_session(
        app_name=APP_NAME,
        user_id=USER_ID,
        state={}
    )
    session_id = session.id
    
    while True:
        # Use asyncio.to_thread to run the blocking input() call without freezing the event loop
        try:
            user_input = await asyncio.to_thread(input, "You: ")
        except (EOFError, KeyboardInterrupt):
            print("\nGoodbye!")
            break
            
        if user_input.lower() == 'quit': 
            break
            
        print("Agent: ", end="", flush=True)
        
        # Create a content message in ADK format
        content = types.Content(
            role='user', 
            parts=[types.Part(text=user_input)]
        )
        
        # THE FIX IS HERE: Use a regular 'for' loop, not 'async for'
        # The runner.run() method returns a standard synchronous generator.
        try:
            for event in runner.run(
                user_id=USER_ID,
                session_id=session_id,
                new_message=content
            ):
                # Print the agent's response as it streams
                if event.content and event.content.parts:
                    for part in event.content.parts:
                        if hasattr(part, 'text') and part.text:
                            print(part.text, end="", flush=True)
            print() # Newline after the full response
        except Exception as e:
            print(f"\nAn error occurred: {e}")
            break

# --- Step 6: Main Execution Block ---
if __name__ == "__main__":
    # Suppress LiteLLM logging errors
    logging.getLogger("litellm").setLevel(logging.CRITICAL)
    logging.getLogger("litellm.logging_utils").setLevel(logging.CRITICAL)
    
    # Redirect stderr to suppress the error messages
    original_stderr = sys.stderr
    sys.stderr = open(os.devnull, 'w')
    
    try:
        # Start the async application
        asyncio.run(run_donna_assistant())
    finally:
        # Restore stderr
        sys.stderr = original_stderr