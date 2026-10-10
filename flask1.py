import os
import threading
from datetime import datetime

from flask import Flask, render_template, jsonify, request
from agent1 import run_conversation
from audio1 import text_audio
import background
import config
import mailer
import scheduled_mail
import task_tracker
app = Flask(__name__)
# Gmail rejects messages over 25 MB, and attachments grow about 33% when encoded
app.config['MAX_CONTENT_LENGTH'] = 18 * 1024 * 1024

@app.route('/process_message', methods=['POST'])
def process_message_func1():
    msg=request.json['message']
    print(msg)
    resp = run_conversation(msg)
    # text_audio(resp)
    return jsonify({"response": resp })

@app.route('/polish_email', methods=['POST'])
def polish_email_route():
    polished = mailer.polish_email(request.json.get('subject', ''), request.json.get('body', ''))
    if not polished:
        return jsonify({"error": "Could not polish the email"}), 500
    return jsonify(polished)

@app.route('/send_email', methods=['POST'])
def send_email_route():
    to = request.form.get('to', '').strip()
    cc = request.form.get('cc', '').strip()
    bcc = request.form.get('bcc', '').strip()
    subject = request.form.get('subject', '')
    body = request.form.get('body', '')
    html = request.form.get('html', '') or None
    send_at = request.form.get('send_at', '').strip()
    if not to:
        return jsonify({"response": "Please add at least one recipient in To"}), 400
    attachments = [
        (f.filename, f.read(), f.mimetype)
        for f in request.files.getlist('attachments') if f.filename
    ]

    if send_at:
        try:
            when = datetime.strptime(send_at, "%Y-%m-%dT%H:%M")
        except ValueError:
            return jsonify({"response": "Invalid schedule time"}), 400
        if when <= datetime.now():
            return jsonify({"response": "Pick a time in the future to schedule the email"}), 400
        for value in (to, cc, bcc):
            invalid = mailer.parse_addresses(value)[1]
            if invalid:
                return jsonify({"response": "Invalid address " + ", ".join(invalid)}), 400
        scheduled_mail.schedule(when, to, cc, bcc, subject, body, html, attachments)
        return jsonify({"response": f"Email scheduled for {when.strftime('%d %b %Y, %I:%M %p')}", "scheduled": True})

    return jsonify({"response": mailer.send_email(to, subject, body, attachments, cc=cc, bcc=bcc, html=html)})

@app.route('/scheduled')
def scheduled_list():
    return jsonify(scheduled_mail.list_scheduled())

@app.route('/scheduled/<item_id>', methods=['DELETE'])
def scheduled_cancel(item_id):
    return jsonify({"ok": scheduled_mail.cancel(item_id)})

@app.route('/contacts')
def contacts():
    people = mailer.load_contacts()
    for addr in config.task_managers:
        people.setdefault(addr, "")
    return jsonify(sorted(
        [{"email": addr, "name": name} for addr, name in people.items()],
        key=lambda c: (c["name"] or c["email"]).lower()))

@app.route('/tasks')
def tasks_list():
    return jsonify({"tasks": task_tracker.list_tasks(), "status": task_tracker.get_status()})

@app.route('/tasks/check', methods=['POST'])
def tasks_check():
    if task_tracker.status["running"]:
        return jsonify({"response": "A check is already running"})
    threading.Thread(target=task_tracker.run_sync, daemon=True).start()
    return jsonify({"response": "Checking your email for new tasks..."})

@app.route('/tasks/<task_id>/status', methods=['POST'])
def tasks_set_status(task_id):
    ok, message = task_tracker.set_status(task_id, request.json.get('status', ''), request.json.get('note'))
    return jsonify({"ok": ok, "response": message}), (200 if ok else 400)

@app.errorhandler(413)
def too_large(e):
    return jsonify({"response": "Attachments are too large. The limit is 18 MB in total."}), 413

@app.route('/')
def hello_world():
    return render_template('index.html')

if __name__ == '__main__':
    # with debug=True Flask runs this file twice; only the reloaded child serves requests
    if os.environ.get("WERKZEUG_RUN_MAIN") == "true":
        background.start()
    app.run(host='0.0.0.0',port=5000,debug=True)
