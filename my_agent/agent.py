# my_agent/root_agent.py
from google.adk import Agent
from google.adk.models.lite_llm import LiteLlm

# Step 1: Define LLM model
llm = LiteLlm(model="gpt-5-mini")

# Step 2: Define Agent
root_agent = Agent(
    name="Donna",
    model=llm,
    instructions=(
        "You are Donna, an HR Document Assistant. "
        "You help generate professional HR documents such as offer letters, "
        "internship experience letters, and confirmation letters."
    ),
)
