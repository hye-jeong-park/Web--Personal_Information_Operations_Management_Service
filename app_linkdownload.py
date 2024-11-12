## 저장할 파일 선택하고(구현 완료), 링크 통해서 파일 다운로드하는 방법(추가 구현 필요) ##

from flask import Flask, render_template, request, session, redirect, url_for
import os
from werkzeug.utils import secure_filename
import sys

app = Flask(__name__)
app.secret_key = 'your_secret_key'
UPLOAD_FOLDER = './uploaded_files'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def run_extraction_script(username, password, max_posts, excel_path):
    from scripts.extraction_script import main as first_main
    try:
        excel_file = first_main(username, password, max_posts, excel_path)
        message = '개인정보 신청 이력 저장이 완료되었습니다.'
        return message, excel_file
    except Exception as e:
        message = f'오류 발생: {e}'
        return message, None

def run_delivery_script(username, password, max_posts, excel_path):
    from scripts.delivery_script import main as second_main
    try:
        excel_file = second_main(username, password, max_posts, excel_path)
        message = '개인정보 추출 및 전달이 완료되었습니다.'
        return message, excel_file
    except Exception as e:
        message = f'오류 발생: {e}'
        return message, None

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        session.clear()  # 새로고침 시 초기화

        # 파일 업로드 처리
        file = request.files['excel_file']
        if file:
            filename = secure_filename(file.filename)
            excel_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(excel_path)
        else:
            return render_template('index.html', message="엑셀 파일을 선택하세요.")

        username = request.form.get('username')
        password = request.form.get('password')
        action = request.form.get('action')
        crawl_option = request.form.get('crawl_option')
        max_posts_input = request.form.get('max_posts')

        if crawl_option == 'direct' and max_posts_input:
            max_posts = int(max_posts_input)
        else:
            max_posts = None

        if action == 'save_history':
            message, file_path = run_extraction_script(username, password, max_posts, excel_path)
        elif action == 'extract_and_transfer':
            message, file_path = run_delivery_script(username, password, max_posts, excel_path)
        else:
            message = '알 수 없는 작업입니다.'
            file_path = None

        return render_template('index.html', message=message, file_path=file_path)
    return render_template('index.html')

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True)