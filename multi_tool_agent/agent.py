import datetime
from zoneinfo import ZoneInfo
import os # Import os to check/set environment variables
from google.adk.agents import Agent
# Import the LiteLlm model wrapper from ADK
from google.adk.models.lite_llm import LiteLlm 
from dotenv import load_dotenv
load_dotenv()  # Load environment variables from a .env file if present
# --- Tool Functions (Remain Unchanged) ---

def get_weather(city: str) -> dict:
    """Retrieves the current weather report for a specified city.

    Args:
        city (str): The name of the city for which to retrieve the weather report.

    Returns:
        dict: status and result or error msg.
    """
    if city.lower() == "new york":
        return {
            "status": "success",
            "report": (
                "The weather in New York is sunny with a temperature of 25 degrees"
                " Celsius (77 degrees Fahrenheit)."
            ),
        }
    else:
        return {
            "status": "error",
            "error_message": f"Weather information for '{city}' is not available.",
        }


def get_current_time(city: str) -> dict:
    """Returns the current time in a specified city.

    Args:
        city (str): The name of the city for which to retrieve the current time.

    Returns:
        dict: status and result or error msg.
    """

    if city.lower() == "new york":
        tz_identifier = "America/New_York"
    else:
        return {
            "status": "error",
            "error_message": (
                f"Sorry, I don't have timezone information for {city}."
            ),
        }

    tz = ZoneInfo(tz_identifier)
    now = datetime.datetime.now(tz)
    report = (
        f'The current time in {city} is {now.strftime("%Y-%m-%d %H:%M:%S %Z%z")}'
    )
    return {"status": "success", "report": report}


# ====================================================================
# LITELLM + OPENAI + ADK IMPLEMENTATION
# ====================================================================

# 1. Define the model using the LiteLlm wrapper.
# LiteLLM model string format for OpenAI is typically "openai/<model-name>"
openai_model = LiteLlm(
    model="openai/gpt-4o-mini"  # Use your desired OpenAI model
    # LiteLLM automatically uses the OPENAI_API_KEY environment variable.
)

# 2. Create the Agent using the LiteLlm model instance
root_agent = Agent(
    name="weather_time_agent_openai",
    # Pass the LiteLlm object here
    model=openai_model, 
    description=(
        "Agent to answer questions about the time and weather in a city using an OpenAI model."
    ),
    instruction=(
        "You are a helpful agent who can answer user questions about the time and weather in a city."
        "You must use the provided tools for time and weather questions."
    ),
    tools=[get_weather, get_current_time],
)

print(f"ADK Agent '{root_agent.name}' is now configured to use model: {root_agent.model.model}")