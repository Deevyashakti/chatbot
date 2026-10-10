# this file execute the respective function based on user query
import task1
import claude
from config import model

# Claude expects "input_schema" where Gemini used "parameters"
tools = [
    {"name": d["name"], "description": d["description"], "input_schema": d["parameters"]}
    for d in task1.definations
]

def run_conversation(user_message):
    system_message = """You are Jarvis, a helpful AI assistant. Answer questions and chat directly yourself.
                      Use the tools only when the user asks you to do one of those actions,
                      such as sending an email, checking the weather, finding an IP address,
                      or viewing and updating the tasks their managers assigned by email.""" # first instruction

    data = {
        "model": model,
        "max_tokens": 1024,
        "system": system_message,
        "messages": [{"role": "user", "content": user_message}],
        "tools": tools,
    }
    result, error = claude.create_message(data)
    if error:
        return error

    content = result.get("content", [])
    print("message is", content)
    for block in content:
        if block.get("type") == "tool_use":
            function_name = block.get("name")
            arguments = block.get("input", {})
            print("Claude: function call", function_name, arguments)
            try:
                resp1 = getattr(task1, function_name)(**arguments)
            except Exception as e:
                print("Error: ", e)
                resp1 = "Type again"
            print("Actual response is", resp1)
            return resp1

    print("No function call in response")
    return next((b.get("text") for b in content if b.get("type") == "text"), "")

if __name__ == "__main__":
    user_message = "find ip address of google.com"
    print(run_conversation(user_message))
