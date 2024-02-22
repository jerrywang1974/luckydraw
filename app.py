import csv
import os
import random
import tempfile
import codecs
from datetime import datetime
from flask import Flask, render_template, request, make_response, send_file, Response,send_from_directory
from fpdf import FPDF
from io import StringIO, BytesIO
from flask_sslify import SSLify
import numpy as np
import pandas as pd

app = Flask(__name__, static_url_path='/static')
sslify = SSLify(app)

@app.before_request
def before_request():
    if not request.is_secure:
        url = request.url.replace('http://', 'https://', 1)
        code = 301
        return redirect(url, code=code)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download/sample-employee.xls')
def download1():
    path = 'sample-employee.xls'
    return send_file(path, as_attachment=True)

@app.route('/download/sample-giftlist.xls')
def download2():
    path = 'sample-giftlist.xls'
    return send_file(path, as_attachment=True)

def isNotNaN(num):
        return num == num
@app.route('/randomize', methods=['POST'])
def randomize():
    #Employee List
    xls_file1 = request.files['csv_file1']
    xls_file2 = request.files['csv_file2']
    #num_teams = int(request.form['num_teams'])
    #force_female_balance = 'force_female_balance' in request.form

    # Read the CSV file and parse the data
    employee_data = []
    gift_data = []
    #reader = csv.DictReader(StringIO(csv_file.read().decode('utf-8')))
    reader_employee = pd.read_excel(xls_file1)
    reader_gift = pd.read_excel(xls_file2)
    
    for row1 in reader_employee.values.tolist():
        if (row1[3]==" " or row1[3]=="　" ):
            print(row1[1]," ",row1[2] , "已抽過")
            pass
        if (isNotNaN(row1[3]) or isNotNaN(row1[4])):
            print(row1[1]," ",row1[2] , "已抽過")
            pass
        else:
            print(row1[1]," ",row1[2],"未抽過")
            employee_data.append(row1)
    for row2 in reader_gift.values.tolist():
        if (isNotNaN(row2[2]) or isNotNaN(row2[3]))  :
            print(row2[1]," ",row2[2] , "已抽過")
            pass
        else:
            print(row2[1]," ",row2[2] , "未抽過")
            gift_data.append(row2)


    #print(data) # 參加抽獎人員
    #wait = input("Press Enter to continue.")
    # Randomly shuffle the data
    print("Number of Employee:",len(employee_data))
    random.shuffle(employee_data)
    random.shuffle(gift_data)
    lucky_draw_result=[]
    number_of_times=len(gift_data)
    while gift_data != []:
        if number_of_times > 0:
            gift = random.choice(gift_data)
            emp_win = random.choice(employee_data)
            gift_data.remove(gift)
            employee_data.remove(emp_win)
            emp_win.append(gift[1])
            lucky_draw_result.append(emp_win)

    #print(lucky_draw_result)
    #avg_person=int(len(data)/num_teams)
    #print(avg_person)
    # Divide members into teams
#    wait = input("Press Enter to continue.")

    # Generate the result as a CSV file
    csv_output = StringIO()
    csv_writer = csv.writer(csv_output)
    csv_writer.writerow(['序號','員工序號','員工姓名', '公司別', '抽中獎項'])
    for i, emp_data in enumerate(lucky_draw_result,start=1):
        #for member in number_of_times:
        #print(i,emp_data)
        csv_writer.writerow([i, emp_data[0],emp_data[1],emp_data[2],emp_data[5]])

    csv_output = csv_output.getvalue().encode('big5','replace')
    csv_output = codecs.BOM_UTF8 + csv_output  # Add BOM

    # Save the CSV to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as csv_temp_file:
        csv_temp_file_path = csv_temp_file.name
        #csv_temp_file_path="draw_result.csv"
        #csv_temp_file.write(csv_output.getvalue().encode('utf-8'))
        csv_temp_file.write(csv_output)
        #csv_temp_file.write(csv_output.getvalue().encode('big5','ignore')

       

    return render_template('result.html', teams=lucky_draw_result, csv_file_path=csv_temp_file_path, pdf_file_path=None)

@app.route('/download_csv/<csv_file_path>', methods=['GET', 'POST'])
def download_csv(csv_file_path):
    base_dir = os.getcwd()
    directory = "/home/jerryw/developer/luckydraw"  # replace with your actual directory path
    filename_for_user = datetime.now().strftime("draw_%Y%m%d_%H%M%S.csv")
    csv_file_path = request.args.get('csv_file_path', "")
    csv_file_path = os.path.abspath(csv_file_path)

    if not os.path.commonpath([base_dir, csv_file_path]).startswith(base_dir):
        return "Access denied: You can only download files within the allowed directory."

    if os.path.exists(csv_file_path):
        # get the directory and base file name
        directory, file_name = os.path.split(csv_file_path)

        # send file as an attachment without reading its contents
        response = make_response(
            send_from_directory(directory, path=csv_file_path, as_attachment=True, mimetype='text/csv;charset=big5'))
        response.headers["Content-Disposition"] = f"attachment; filename={filename_for_user}"
        return response
    else:
        return "File does not exist."

# @app.route('/download_pdf')
# def download_pdf():
#     pdf_file_path = request.args.get('pdf_file_path')
#     return send_file(pdf_file_path, as_attachment=True)

if __name__ == '__main__':
    #app.run(debug=True, host='0.0.0.0',port=443,ssl_context=('/etc/letsencrypt/live/monitor.jerryw.org/fullchain.pem', '/etc/letsencrypt/live/monitor.jerryw.org/privkey.pem'))
    app.run(debug=True, host='::',port=443,ssl_context=('/etc/letsencrypt/live/monitor.jerryw.org/fullchain.pem', '/etc/letsencrypt/live/monitor.jerryw.org/privkey.pem'))
