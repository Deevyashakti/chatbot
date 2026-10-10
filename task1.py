import requests
from config import headers
import socket
import mailer
import task_tracker

def get_ip(host):
    try:
        ips = sorted({info[4][0] for info in socket.getaddrinfo(host, None)})
        result = f"IP address of {host}: {', '.join(ips)}"
    except Exception as e:
        print(e)
        result = f"Error in find the IP, {e}"
    return result

def temp_room(room):
    result = f"{room} Temperature = 20, Humidity = 50"
    return result

def temp_city(city):
    url = "https://yahoo-weather5.p.rapidapi.com/weather"

    querystring = {"location":city,"format":"json","u":"f"}

    response = requests.get(url, headers=headers, params=querystring)

    d1 = response.json()
    d1 = d1.get("current_observation")
    hum = d1.get("atmosphere").get("humidity")
    temp = d1.get("condition").get("temperature")
    temp = round((temp-32)*5/9,2)
    return (f"humidity:  {hum}, Temperature in C: {temp}")

def send_email(to, subject, body, polish=False, cc="", bcc=""):
    if polish:
        polished = mailer.polish_email(subject, body)
        if polished:
            subject, body = polished["subject"], polished["body"]
    return mailer.send_email(to, subject, body, cc=cc, bcc=bcc)

def list_tasks(status=""):
    tasks = task_tracker.list_tasks()
    if status:
        tasks = [t for t in tasks if t["status"].lower() == status.lower()]
    if not tasks:
        return f"No {status.lower() + ' ' if status else ''}tasks found."
    lines = []
    for t in tasks[:30]:
        line = f"{t['id']} [{t['status'] or 'Not Started'}] {t['task']}"
        if t["deadline"]:
            line += f" (due {t['deadline']})"
        lines.append(line)
    more = f"\n...and {len(tasks) - 30} more" if len(tasks) > 30 else ""
    return f"{len(tasks)} task(s):\n" + "\n".join(lines) + more

def update_task_status(task_id, status, note=""):
    return task_tracker.set_status(task_id, status, note)[1]

def check_task_emails():
    return task_tracker.run_sync()

definations = [
    {
        "name":"list_tasks",
        "description": "list the tasks assigned to the user by their managers, optionally filtered by status",
        "parameters":
            {
                "type":"object",
                "properties":{
                    "status" : {
                        "type":"string",
                        "enum":["", "Not Started", "In Progress", "Completed"],
                        "description":"filter by status, empty for all tasks"
                    }
                }
            }
    },
    {
        "name":"update_task_status",
        "description": "change a task's status, e.g. mark T-004 in progress or completed. Completing a task automatically replies to the manager's email",
        "parameters":
            {
                "type":"object",
                "properties":{
                    "task_id" : {"type":"string", "description":"task ID such as T-004"},
                    "status" : {"type":"string", "enum":["Not Started", "In Progress", "Completed"]},
                    "note" : {"type":"string", "description":"optional note to include in the completion reply"}
                },
                "required": ["task_id", "status"]
            }
    },
    {
        "name":"check_task_emails",
        "description": "check the inbox right now for new task emails from managers",
        "parameters": {"type":"object", "properties":{}}
    },
    {
        "name":"send_email",  # name of the function to be called
        "description": "send an email to someone",
        "parameters":
            {
                "type":"object",
                "properties":{
                    "to" : {
                        "type":"string",
                        "description":"recipient email address"
                    },
                    "subject" : {
                        "type":"string",
                        "description":"subject of the email"
                    },
                    "body" : {
                        "type":"string",
                        "description":"body text of the email, exactly as the user wants it"
                    },
                    "polish" : {
                        "type":"boolean",
                        "description":"true only if the user asks to polish, improve or make the email professional"
                    },
                    "cc" : {"type":"string", "description":"comma-separated Cc addresses, if any"},
                    "bcc" : {"type":"string", "description":"comma-separated Bcc addresses, if any"}
                },
                "required": ["to", "subject", "body"]
            }
    },
    {
        "name":"temp_city",  # name of the function to be called
        "description": "find weather, temperature of a city",
        "parameters":
            {
                "type":"object",
                "properties":{
                    "city" : {                 # Argument for function temp_city
                        "type":"string",
                        "description":"city to find weather"
                    }
                }
            }
    },
    {
        "name":"temp_room",  # name of the function to be called
        "description": "find temperature of my room or my home",
        "parameters":
            {
                "type":"object",
                "properties":{
                    "room" : {                 # Argument for function temp_city
                        "type":"string",
                        "description":"room or home"
                    }
                }
            }
    },
    {
        "name":"get_ip",  # name of the function to be called
        "description": "find ip address of given url or domain name",
        "parameters":
            {
                "type":"object",
                "properties":{
                    "host" : {                 # Argument for function temp_city
                        "type":"string",
                        "description":"get url or domain name"
                    }
                }
            }
    }
]

if __name__ == "__main__":
    print(temp_city("Hyderabad"))