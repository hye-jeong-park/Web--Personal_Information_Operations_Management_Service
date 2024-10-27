from flask import Flask, render_template, request
import threading
import os

app = Flask(__name__)

# 스크립트가 저장된 디렉토리
SCRIPT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'scripts')

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        action = request.form['action']

        if action == 'save_history':
            message, file_path = run_extraction_script(username, password)
        elif action == 'extract_and_transfer':
            message, file_path = run_delivery_script(username, password)
        else:
            message = '알 수 없는 동작입니다.'
            file_path = None

        return render_template('index.html', message=message, file_path=file_path)
    else:
        return render_template('index.html')