from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash
from werkzeug.utils import secure_filename
import os
import logging
import uuid
from create_bible_format import ChangeBibleFormat
from create_song_form import ChangeLyrics
from google.cloud import storage

# Flask 애플리케이션 초기화
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', '3dce87f15d8d20bd3e3b80c6fe349bb997497a8899c4cd780076a7be653739d7')  # 환경 변수에서 가져오기

# Cloud Storage 설정
GCS_BUCKET = os.environ.get('GCS_BUCKET')  # 환경 변수로 설정
storage_client = storage.Client()
bucket = storage_client.bucket(GCS_BUCKET)

# 폴더 설정
UPLOAD_FOLDER = 'uploads'  # 현재는 사용하지 않음, 필요 시 유지
BIBLE_PPT_FOLDER = 'bible_ppts'
OUTPUT_FOLDER = 'output'
ALLOWED_TEXT_EXTENSIONS = {'txt'}
ALLOWED_PPT_EXTENSIONS = {'pptx'}

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
def get_version_sets():
    version_sets = []
    if os.path.isdir(BIBLE_PPT_FOLDER):
        for folder in os.listdir(BIBLE_PPT_FOLDER):
            full_path = os.path.join(BIBLE_PPT_FOLDER, folder)
            if os.path.isdir(full_path):
                version_sets.append(folder)
                logging.info(f"Found version set: {folder}")
    else:
        logging.error(f"BIBLE_PPT_FOLDER does not exist: {BIBLE_PPT_FOLDER}")
    logging.info(f"Total version sets found: {len(version_sets)}")
    return version_sets

# 파일을 GCS에 업로드하는 함수
def upload_to_gcs(source_file_path, destination_blob_name):
    try:
        blob = bucket.blob(destination_blob_name)
        blob.upload_from_filename(source_file_path)
        blob.make_public()  # 공개적으로 접근 가능하게 설정 (필요 시)
        return blob.public_url
    except Exception as e:
            logging.error(f"Google Cloud Storage 업로드 실패: {e}")
            raise e

# 루트 라우트: 파일 업로드 및 입력 폼
@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')

# ChangeBibleFormat 라우트
# @app.route('/create_bible_format', methods=['GET'])
# def create_bible_format_menu():
#     return render_template('create_bible_format.html')
@app.route('/create_bible_format', methods=['GET', 'POST'])
def create_bible_format_route():
    if request.method == 'POST':
        # 폼 데이터 가져오기
        text_file = request.files.get('text_file')
        ppt_title = request.form.get('ppt_title')
        version_set = request.form.get('version_set')

        # 유효성 검사
        if not text_file or text_file.filename == '':
            flash('텍스트 파일을 선택해주세요.', 'warning')
            return redirect(request.url)
        if not allowed_file(text_file.filename, ALLOWED_TEXT_EXTENSIONS):
            flash('허용되지 않은 파일 형식입니다. .txt 파일만 업로드할 수 있습니다.', 'warning')
            return redirect(request.url)
        if not version_set:
            flash('버전을 선택해주세요.', 'warning')
            return redirect(request.url)
        if not ppt_title:
            flash('PPT 제목을 입력해주세요.', 'warning')
            return redirect(request.url)

        # 고유한 폴더 이름 생성 (UUID 사용) - 필요 시 사용
        unique_id = str(uuid.uuid4())

        # 업로드된 텍스트 파일 저장
        text_filename = secure_filename(text_file.filename)
        local_text_path = os.path.join('/tmp', f"{unique_id}_{text_filename}")
        text_file.save(local_text_path)

        try:
            # 선택한 번역본 폴더 경로 설정
            selected_version_path = os.path.join(app.config['BIBLE_PPT_FOLDER'], version_set)
            if not os.path.isdir(selected_version_path):
                flash('선택한 버전의 파일이 존재하지 않습니다.', 'danger')
                return redirect(request.url)

            logging.info(f"Selected version path: {selected_version_path}")

            # ChangeBibleFormat 인스턴스 생성 및 PPT 병합 실행
            cbf = ChangeBibleFormat(
                text_file_path=local_text_path,
                bible_path=selected_version_path,
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
    version_sets = get_version_sets()
    logging.info(f"versions passed to template: {version_sets}")
    return render_template('create_bible_format.html', version_sets=version_sets)

# ChangeLyrics 메인 메뉴 라우트
@app.route('/change_lyrics', methods=['GET'])
def change_lyrics_menu():
    return render_template('change_lyrics.html')

# ChangeLyrics 작업별 라우트
@app.route('/change_lyrics/<operation>', methods=['GET', 'POST'])
def change_lyrics_route(operation):
    allowed_operations = [
        'change_lyrics',
        'change_only_korean_lyrics',
        'change_only_english_lyrics',
        'create_lyrics',
        'create_only_korean_lyrics',
        'create_only_english_lyrics',
        'create_seoul_form'
    ]

    if operation not in allowed_operations:
        flash('알 수 없는 작업이 선택되었습니다.', 'danger')
        return redirect(url_for('change_lyrics_menu'))

    if request.method == 'POST':
        # 폼 데이터 가져오기
        text_file = request.files.get('text_file')
        ppt_file = request.files.get('ppt_file')  # PPT 파일 업로드 필드 추가
        ppt_title = request.form.get('ppt_title')

        # 유효성 검사
        if not text_file or text_file.filename == '':
            flash('텍스트 파일을 선택해주세요.', 'warning')
            return redirect(request.url)
        if not allowed_file(text_file.filename, ALLOWED_TEXT_EXTENSIONS):
            flash('허용되지 않은 파일 형식입니다. .txt 파일만 업로드할 수 있습니다.', 'warning')
            return redirect(request.url)

        # 특정 작업에 따라 PPT 파일이 필요한지 확인
        operations_requiring_ppt = [
            'change_lyrics',
            'change_only_korean_lyrics',
            'change_only_english_lyrics'
        ]
        if operation in operations_requiring_ppt:
            if not ppt_file or ppt_file.filename == '':
                flash('기존 PPT 파일을 업로드해주세요.', 'warning')
                return redirect(request.url)
            if not allowed_file(ppt_file.filename, ALLOWED_PPT_EXTENSIONS):
                flash('허용되지 않은 파일 형식입니다. .pptx 파일만 업로드할 수 있습니다.', 'warning')
                return redirect(request.url)

        if not ppt_title:
            flash('PPT 제목을 입력해주세요.', 'warning')
            return redirect(request.url)

        # 고유한 폴더 이름 생성 (UUID 사용)
        unique_id = str(uuid.uuid4())

        # 업로드된 텍스트 파일 저장
        text_filename = secure_filename(text_file.filename)
        local_text_path = os.path.join('/tmp', f"{unique_id}_{text_filename}")
        text_file.save(local_text_path)

        # 업로드된 PPT 파일 저장 (필요 시)
        if operation in operations_requiring_ppt:
            ppt_filename = secure_filename(ppt_file.filename)
            local_ppt_path = os.path.join('/tmp', f"{unique_id}_{ppt_filename}")
            ppt_file.save(local_ppt_path)
        else:
            local_ppt_path = None  # 새로 생성할 경우 기존 PPT가 없으므로 None

        try:
            # ChangeLyrics 인스턴스 초기화
            cl = ChangeLyrics(
                text_file_path=local_text_path,
                ppt_path=local_ppt_path,  # 기존 PPT 파일 경로 또는 None
                save_path='/tmp',
                ppt_title=ppt_title
            )
            # 선택된 작업 메서드 호출
            getattr(cl, operation)()

            # 생성된 PPT 파일 경로
            ppt_output_filename = f"{ppt_title}.pptx"
            local_ppt_output_path = os.path.join('/tmp', ppt_output_filename)

            if os.path.exists(local_ppt_output_path):
                # PPT 파일을 GCS에 업로드
                gcs_blob_name = f"output/{ppt_output_filename}"
                gcs_url = upload_to_gcs(local_ppt_output_path, gcs_blob_name)

                flash('PPT 파일이 성공적으로 생성되었습니다.', 'success')
                return redirect(gcs_url)  # GCS URL로 리디렉션하여 다운로드
            else:
                flash('PPT 파일 생성에 실패했습니다.', 'danger')
                return redirect(request.url)

        except Exception as e:
            logging.error(f"에러 발생: {e}")
            flash(f"에러 발생: {e}", 'danger')
            return redirect(request.url)

    # GET 요청 시 작업에 따라 필요한 정보만 제공
    return render_template('change_lyrics_operations.html', operation=operation)


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
