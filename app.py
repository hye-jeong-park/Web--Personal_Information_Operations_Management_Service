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
    
def run_extraction_script(username, password):
    # "개인정보 추출 신청" 스크립트 실행
    from scripts.extraction_script import main as first_main
    try:
        excel_file = first_main(username, password)
        message = '개인정보 신청 이력 저장이 완료되었습니다.'
        return message, excel_file
    except Exception as e:
        message = f'오류 발생: {e}'
        return message, None

def run_delivery_script(username, password):
    # "개인정보 추출 및 전달" 스크립트 실행
    from scripts.delivery_script import main as second_main
    try:
        excel_file = second_main(username, password)
        message = '개인정보 추출 및 전달이 완료되었습니다.'
        return message, excel_file
    except Exception as e:
        message = f'오류 발생: {e}'
        return message, None

if __name__ == '__main__':
    app.run(debug=False)