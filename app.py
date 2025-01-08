from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash
from werkzeug.utils import secure_filename
import os
import logging
import uuid
from create_bible_format import ChangeBibleFormat  # ChangeBibleFormat 클래스가 포함된 모듈
from google.cloud import storage

# Flask 애플리케이션 초기화
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY')  # 환경 변수에서 가져오기

# Cloud Storage 설정
GCS_BUCKET = os.environ.get('GCS_BUCKET')  # 환경 변수로 설정
storage_client = storage.Client()
bucket = storage_client.bucket(GCS_BUCKET)

# 폴더 설정
UPLOAD_FOLDER = 'uploads'  # 현재는 사용하지 않음, 필요 시 유지
BIBLE_PPT_FOLDER = 'bible_ppts'
OUTPUT_FOLDER = 'output'
ALLOWED_TEXT_EXTENSIONS = {'txt'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['BIBLE_PPT_FOLDER'] = BIBLE_PPT_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 최대 업로드 파일 크기 설정 (예: 100MB)

# 로그 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 업로드 가능한 파일인지 확인하는 함수
def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

# 번역본 목록 가져오기 함수
def get_translation_sets():
    translation_sets = []
    if os.path.isdir(BIBLE_PPT_FOLDER):
        for folder in os.listdir(BIBLE_PPT_FOLDER):
            full_path = os.path.join(BIBLE_PPT_FOLDER, folder)
            if os.path.isdir(full_path):
                translation_sets.append(folder)
                logging.info(f"Found translation set: {folder}")
    else:
        logging.error(f"BIBLE_PPT_FOLDER does not exist: {BIBLE_PPT_FOLDER}")
    logging.info(f"Total translation sets found: {len(translation_sets)}")
    return translation_sets

# 파일을 GCS에 업로드하는 함수
def upload_to_gcs(source_file_path, destination_blob_name):
    blob = bucket.blob(destination_blob_name)
    blob.upload_from_filename(source_file_path)
    blob.make_public()  # 공개적으로 접근 가능하게 설정 (필요 시)
    return blob.public_url

# 루트 라우트: 파일 업로드 및 입력 폼
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 폼 데이터 가져오기
        text_file = request.files.get('text_file')
        ppt_title = request.form.get('ppt_title')
        translation_set = request.form.get('translation_set')

        # 유효성 검사
        if not text_file or text_file.filename == '':
            flash('텍스트 파일을 선택해주세요.', 'warning')
            return redirect(request.url)
        if not allowed_file(text_file.filename, ALLOWED_TEXT_EXTENSIONS):
            flash('허용되지 않은 파일 형식입니다. .txt 파일만 업로드할 수 있습니다.', 'warning')
            return redirect(request.url)
        if not translation_set:
            flash('번역본을 선택해주세요.', 'warning')
            return redirect(request.url)
        if not ppt_title:
            flash('PPT 제목을 입력해주세요.', 'warning')
            return redirect(request.url)

        # 고유한 폴더 이름 생성 (UUID 사용) - 필요 시 사용
        unique_id = str(uuid.uuid4())

        # 업로드된 텍스트 파일 저장
        text_filename = secure_filename(text_file.filename)
        text_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{unique_id}_{text_filename}")
        local_text_path = os.path.join('/tmp', f"{unique_id}_{text_filename}")
        text_file.save(local_text_path)

        try:
            # 선택한 번역본 폴더 경로 설정
            selected_translation_path = os.path.join(app.config['BIBLE_PPT_FOLDER'], translation_set)
            if not os.path.isdir(selected_translation_path):
                flash('선택한 번역본 폴더가 존재하지 않습니다.', 'danger')
                return redirect(request.url)

            logging.info(f"Selected translation path: {selected_translation_path}")

            # ChangeBibleFormat 인스턴스 생성 및 PPT 병합 실행
            cbf = ChangeBibleFormat(
                text_file_path=local_text_path,
                bible_path=selected_translation_path,
                output_path='/tmp',
                ppt_title=ppt_title
            )
            cbf.create_ppt_file()

            # 생성된 PPT 파일 경로
            ppt_filename = f"{ppt_title}.pptx"
            local_ppt_path = os.path.join('/tmp', ppt_filename)

            if os.path.exists(local_ppt_path):
                # PPT 파일을 GCS에 업로드
                gcs_blob_name = f"output/{ppt_filename}"
                gcs_url = upload_to_gcs(local_ppt_path, gcs_blob_name)

                flash('PPT 파일이 성공적으로 생성되었습니다.', 'success')
                return redirect(gcs_url)  # GCS URL로 리디렉션하여 다운로드
            else:
                flash('PPT 파일 생성에 실패했습니다.', 'danger')
                return redirect(request.url)

        except Exception as e:
            logging.error(f"에러 발생: {e}")
            flash(f"에러 발생: {e}", 'danger')
            return redirect(request.url)

    # GET 요청 시 번역본 목록 가져오기
    translation_sets = get_translation_sets()
    logging.info(f"Translations passed to template: {translation_sets}")
    return render_template('index.html', translation_sets=translation_sets)

# PPT 파일 다운로드 라우트 (필요 시)
@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(directory=app.config['OUTPUT_FOLDER'], path=filename, as_attachment=True)

if __name__ == '__main__':
    # 폴더가 없으면 생성
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(BIBLE_PPT_FOLDER, exist_ok=True)
    os.makedirs('/tmp', exist_ok=True)
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5050)), debug=False)
