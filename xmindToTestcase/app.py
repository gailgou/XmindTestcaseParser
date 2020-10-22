from flask import Flask, render_template, request, send_from_directory, session, make_response
from src.main import *
import os
import time
import urllib

upload_path = 'upload'
app = Flask(__name__)
app.secret_key = 'bytedance'


@app.route('/index', methods=['GET', 'POST'])
def index():
    session['tag'] = False
    if request.method == 'POST':
        res = {
            'error': '',
            'file_url': '',
            'success_msg': ''
        }
        del_files()

        xmind_file_obj = request.files['xmindFile']
        xmind_file_name = xmind_file_obj.filename
        creator = request.form.get('creator')
        is_need_result = request.form.get('needResult')

        if not xmind_file_name.endswith('.xmind'):
            res['error'] = '上传的 xmind 文件不正确！'
            return render_template('index.html', res=res)

        save_path = os.path.join(upload_path, xmind_file_name)
        xmind_file_obj.save(save_path)
        excel_file_name = xmind_file_name.rsplit(".", 1)[0] + '.xls'

        param = {
            'creator': creator,
            'is_need_result': is_need_result,
            'xmind_path': save_path,
            'excel_path': os.path.join(upload_path, excel_file_name)
        }

        if 'all' in request.form:
            get_all_xmind_case(param)
            res['file_url'] = os.path.join('/download/', excel_file_name)
            res['success_msg'] = xmind_file_name + ' 转换成功，点击下载用例！'
            return render_template('index.html', res=res)
        elif 'develop_only' in request.form:
            get_develop_xmind_case(param)
            res['file_url'] = os.path.join('/download/', excel_file_name)
            res['success_msg'] = xmind_file_name + ' 转换成功，点击下载用例！'
            return render_template('index.html', res=res)
        elif 'get_num_rate' in request.form:
            res = get_num_and_rate(param)
            return render_template('result.html', res=res)

    return render_template('index.html', res={})


@app.route('/download/<filename>', methods=['GET'])
def download(filename):
    file_name = urllib.request.unquote(filename, encoding='utf-8', errors='replace')
    excel_file_path = os.path.join(upload_path, file_name)
    if request.method == "GET":
        if os.path.isfile(excel_file_path):
            return send_from_directory(upload_path, file_name, as_attachment=True)


def del_files():
    for file_name in os.listdir(upload_path):
        del_files_path = os.path.join(upload_path, file_name)
        # 获取文件的创建日期
        create_time = time.localtime(os.stat(del_files_path).st_ctime)
        create_date = time.strftime("%Y-%m-%d", create_time)
        # 隔天删除
        if create_date < time.strftime("%Y-%m-%d"):
            os.remove(del_files_path)


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)
