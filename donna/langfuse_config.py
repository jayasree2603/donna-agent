"""
Langfuse integration module for Donna HR Assistant.
This module provides comprehensive tracing for LLM requests, responses, and tool calls.
"""

import os
import functools
from typing import Any, Callable, Dict, Optional, List
from dotenv import load_dotenv
from datetime import datetime

from langfuse import Langfuse, observe
from google.genai import types

load_dotenv()

# ============================================================================
# LANGFUSE CLIENT INITIALIZATION
# ============================================================================

langfuse_client = Langfuse(
    public_key=os.getenv("LANGFUSE_PUBLIC_KEY"),
    secret_key=os.getenv("LANGFUSE_SECRET_KEY"),
    host=os.getenv("LANGFUSE_HOST", "https://cloud.langfuse.com"),
    
)

# ============================================================================
# LLM REQUEST/RESPONSE TRACING
# ============================================================================

class LLMTracer:
    """Tracer for LLM requests and responses."""

    def __init__(self, model_name: str = "gpt-4o-mini"):
        self.model_name = model_name
        self.current_generation = None

    def start_generation(self, user_input: str, session_id: str, metadata: Optional[Dict[str, Any]] = None):
        self.current_generation = langfuse_client.generation(
            name="llm_generation",
            model=self.model_name,
            input=user_input,
            metadata={
                "session_id": session_id,
                "timestamp": datetime.now().isoformat(),
                **(metadata or {})
            }
        )
        return self.current_generation

    def log_llm_call(self, prompt: str, response: str, tool_calls: Optional[List[Dict]] = None, metadata: Optional[Dict[str, Any]] = None):
        generation_metadata = {
            "timestamp": datetime.now().isoformat(),
            "has_tool_calls": bool(tool_calls),
            "tool_count": len(tool_calls) if tool_calls else 0,
            **(metadata or {})
        }
        if tool_calls:
            generation_metadata["tool_calls"] = tool_calls

        langfuse_client.generation(
            name="llm_request_response",
            model=self.model_name,
            input=prompt,
            output=response,
            metadata=generation_metadata
        )

    def end_generation(self, output: str, usage: Optional[Dict[str, int]] = None, metadata: Optional[Dict[str, Any]] = None):
        if self.current_generation:
            update_data = {"output": output, "metadata": metadata or {}}
            if usage:
                update_data["usage"] = usage
            self.current_generation.update(**update_data)
            self.current_generation = None


# ============================================================================
# TOOL CALL TRACING
# ============================================================================

class ToolCallTracer:
    """Tracer for tool/function calls."""

    @staticmethod
    def log_tool_call(tool_name: str, input_args: Dict[str, Any], output: Any, success: bool = True, error: Optional[str] = None, execution_time: Optional[float] = None):
        langfuse_client.span(
            name=f"tool_{tool_name}",
            input=input_args,
            output=output,
            metadata={
                "tool_name": tool_name,
                "success": success,
                "error": error,
                "execution_time_ms": execution_time * 1000 if execution_time else None,
                "timestamp": datetime.now().isoformat()
            }
        )

    @staticmethod
    def trace_tool_execution(tool_name: str):
        """Decorator to automatically trace tool execution."""
        def decorator(func: Callable) -> Callable:
            @observe(name=f"tool_{tool_name}")
            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                import time
                start_time = time.time()
                try:
                    result = func(*args, **kwargs)
                    execution_time = time.time() - start_time
                    return result
                except Exception as e:
                    raise e
            return wrapper
        return decorator


# ============================================================================
# RUNNER EVENT TRACING
# ============================================================================

class RunnerEventTracer:
    """Tracer for Google ADK Runner events."""

    def __init__(self, session_id: str, user_id: str):
        self.session_id = session_id
        self.user_id = user_id
        self.current_trace = None
        self.tool_calls_in_turn = []
        self.llm_tracer = LLMTracer()

    def start_turn(self, user_input: str):
        self.current_trace = langfuse_client.trace(
            name="conversation_turn",
            user_id=self.user_id,
            session_id=self.session_id,
            input={"user_message": user_input},
            metadata={"timestamp": datetime.now().isoformat()}
        )
        self.tool_calls_in_turn = []
        self.llm_tracer.start_generation(user_input, self.session_id)
        return self.current_trace

    def log_tool_call(self, tool_name: str, arguments: Dict, result: Any):
        self.tool_calls_in_turn.append({
            "tool_name": tool_name,
            "arguments": arguments,
            "result": result,
            "timestamp": datetime.now().isoformat()
        })
        langfuse_client.span(
            name=f"tool_call_{tool_name}",
            trace_id=self.current_trace.id if self.current_trace else None,
            input=arguments,
            output=result,
            metadata={"tool_name": tool_name, "success": result.get("success") if isinstance(result, dict) else True}
        )

    def log_llm_response(self, response_text: str, metadata: Optional[Dict] = None):
        if self.current_trace:
            langfuse_client.generation(
                name="llm_response",
                trace_id=self.current_trace.id,
                model="gpt-4o-mini",
                output=response_text,
                metadata={
                    "tool_calls_count": len(self.tool_calls_in_turn),
                    "has_tool_calls": len(self.tool_calls_in_turn) > 0,
                    **(metadata or {})
                }
            )

    def end_turn(self, assistant_response: str):
        if self.current_trace:
            self.current_trace.update(
                output={"assistant_response": assistant_response},
                metadata={
                    "tool_calls": self.tool_calls_in_turn,
                    "tool_count": len(self.tool_calls_in_turn),
                    "completed_at": datetime.now().isoformat()
                }
            )
            self.llm_tracer.end_generation(assistant_response)
        self.current_trace = None
        self.tool_calls_in_turn = []


# ============================================================================
# WRAPPER FUNCTIONS
# ============================================================================

def trace_validation(func: Callable) -> Callable:
    @observe(name=f"validation_{func.__name__}")
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        return func(*args, **kwargs)
    return wrapper

def trace_tool(func: Callable) -> Callable:
    return ToolCallTracer.trace_tool_execution(func.__name__)(func)

def trace_conversation_turn(session_id: str, user_id: str):
    def decorator(func: Callable) -> Callable:
        @observe(name="conversation_turn")
        @functools.wraps(func)
        async def wrapper(*args, **kwargs):
            return await func(*args, **kwargs)
        return wrapper
    return decorator


# ============================================================================
# SESSION TRACING
# ============================================================================

class LangfuseSessionTracer:
    """Context manager for tracing entire conversation sessions."""

    def __init__(self, app_name: str, user_id: str, session_id: str):
        self.app_name = app_name
        self.user_id = user_id
        self.session_id = session_id
        self.trace = None
        self.conversation_count = 0
        self.event_tracer = None

    def __enter__(self):
        self.trace = langfuse_client.trace(
            name="donna_session",
            user_id=self.user_id,
            session_id=self.session_id,
            metadata={
                "app_name": self.app_name,
                "session_start": datetime.now().isoformat()
            }
        )
        self.event_tracer = RunnerEventTracer(self.session_id, self.user_id)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.trace:
            status = "error" if exc_type else "completed"
            self.trace.update(
                metadata={
                    "conversation_turns": self.conversation_count,
                    "status": status,
                    "session_end": datetime.now().isoformat(),
                    "error": str(exc_val) if exc_val else None
                }
            )
        langfuse_client.flush()

    def increment_turn(self):
        self.conversation_count += 1

    def get_event_tracer(self) -> RunnerEventTracer:
        return self.event_tracer

    def log_event(self, event_name: str, metadata: Optional[Dict[str, Any]] = None):
        if self.trace:
            langfuse_client.span(
                name=event_name,
                trace_id=self.trace.id,
                metadata={
                    "timestamp": datetime.now().isoformat(),
                    **(metadata or {})
                }
            )


# ============================================================================
# ERROR TRACKING
# ============================================================================

def track_error(error: Exception, context: Optional[Dict[str, Any]] = None):
    langfuse_client.trace(
        name="donna_error",
        metadata={
            "error_type": type(error).__name__,
            "error_message": str(error),
            "timestamp": datetime.now().isoformat(),
            "context": context or {}
        }
    )
    langfuse_client.flush()


# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def flush_traces():
    langfuse_client.flush()

def create_generation_trace(name: str, model: str, input_text: str, output_text: str, metadata: Optional[Dict[str, Any]] = None):
    langfuse_client.generation(
        name=name,
        model=model,
        input=input_text,
        output=output_text,
        metadata={"timestamp": datetime.now().isoformat(), **(metadata or {})}
    )


# ============================================================================
# EXPORT
# ============================================================================

__all__ = [
    'langfuse_client',
    'observe',
    'trace_validation',
    'trace_tool',
    'trace_conversation_turn',
    'LangfuseSessionTracer',
    'LLMTracer',
    'ToolCallTracer',
    'RunnerEventTracer',
    'track_error',
    'flush_traces',
    'create_generation_trace'
]
