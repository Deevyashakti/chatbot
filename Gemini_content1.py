from config import model
import claude
from micToText import mic1


def chat1(chat):
    system_message = "You are an AI bot, your name is Jarvis. " # first instruction
    data = {
        "model": model,
        "max_tokens": 1024,
        "system": system_message,
        "messages": [{"role": "user", "content": chat}],
    }
    t1, error = claude.create_message(data)
    if error:
        print(error)
        return
    t2 = t1.get("content")[0].get("text")
    print(t2)

# chat = mic1()
chat="who is Deepesh Ahuja"
chat1(chat)
