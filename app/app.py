import os
from flask import Flask, render_template, request, send_file
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
            map = 'templates/map.html'
            draw_map(filename=os.path.join(app.config['UPLOAD_PATH'], filename), output_map=map)
            if 'save' in request.form:
                return send_file(map, as_attachment=True)
            elif 'view' in request.form:
                return render_template('output_map.html')
        except:
            return render_template('process_failed.html')
    else:
        return render_template('file_error.html')

if __name__ == '__main__':
    app.run(debug=True)