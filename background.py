# Background loop: sends scheduled emails and checks the inbox for new tasks.
import threading
import time

import config
import scheduled_mail
import task_tracker

_started = False


def _loop():
    next_task_check = 0
    while True:
        try:
            scheduled_mail.send_due()
        except Exception as e:
            print("Scheduled email error:", e)
        if time.time() >= next_task_check:
            try:
                summary = task_tracker.run_sync()
                print("Task check:", summary)
            except Exception as e:
                print("Task check error:", e)
            next_task_check = time.time() + config.task_check_minutes * 60
        time.sleep(30)


def start():
    global _started
    if _started:
        return
    _started = True
    threading.Thread(target=_loop, name="jarvis-background", daemon=True).start()
