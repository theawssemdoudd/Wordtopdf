from flask import Flask, request, send_file, jsonify
from docx2pdf import convert
import os
import tempfile
import logging
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB

# إعداد السجل
logging.basicConfig(level=logging.INFO)

ALLOWED_EXTENSIONS = {'doc', 'docx'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return send_file('index.html')

@app.route('/convert', methods=['POST'])
def convert_word_to_pdf():
    if 'file' not in request.files:
        return 'لم يتم اختيار ملف', 400
    
    file = request.files['file']
    
    if file.filename == '':
        return 'لم يتم اختيار ملف', 400
    
    if not allowed_file(file.filename):
        return 'نوع الملف غير مدعوم. يرجى استخدام .doc أو .docx', 400

    try:
        # إنشاء ملفات مؤقتة
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_input:
            input_path = temp_input.name
            file.save(input_path)

        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_output:
            output_path = temp_output.name

        # التحويل
        logging.info(f'جاري تحويل: {file.filename}')
        convert(input_path, output_path)
        logging.info('تم التحويل بنجاح')

        # إرجاع الملف المحول
        return send_file(
            output_path,
            as_attachment=True,
            download_name=file.filename.replace('.docx', '.pdf').replace('.doc', '.pdf'),
            mimetype='application/pdf'
        )

    except Exception as e:
        logging.error(f'خطأ في التحويل: {str(e)}')
        return f'خطأ في التحويل: {str(e)}', 500

    finally:
        # تنظيف الملفات المؤقتة
        try:
            if os.path.exists(input_path):
                os.unlink(input_path)
            if os.path.exists(output_path):
                os.unlink(output_path)
        except:
            pass

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
