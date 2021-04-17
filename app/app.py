import os
from flask import Flask, render_template, request, redirect, url_for, abort
from werkzeug.utils import secure_filename
from wmg_tracing_map import *

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024
app.config['UPLOAD_EXTENSIONS'] = ['.xls', '.xlsx']
app.config['UPLOAD_PATH'] = 'data/uploads'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/', methods=['POST'])
def upload_file():
    # return "OK this is a post method"
    uploaded_file = request.files['file']
    filename = secure_filename(uploaded_file.filename)
    if filename != '' and (os.path.splitext(filename)[1]).lower() in app.config['UPLOAD_EXTENSIONS']:
        filename = filename.lower()
        uploaded_file.save(os.path.join(app.config['UPLOAD_PATH'], filename))
        try:
            draw_map(filename=os.path.join(app.config['UPLOAD_PATH'], filename), output_map=os.path.join('templates', 'test_map.html'))
        except:
            return render_template('process_failed.html')
    else:
        return render_template('file_error.html')
    return render_template('test_map.html')

if __name__ == '__main__':
    app.run(debug=True)