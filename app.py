import csv
import random
from flask import Flask, render_template, request, make_response, send_file
from io import StringIO, BytesIO
from fpdf import FPDF
import numpy as np
import tempfile
import pandas as pd

app = Flask(__name__, static_url_path='/static')

@app.route('/')
def index():
    return render_template('index.html')

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
        if row1[3] is not np.nan:
            pass
        else:
            employee_data.append(row1)
    for row2 in reader_gift.values.tolist():
        if (row2[2] or row2[3]) is not np.nan:
            pass
        else:
            gift_data.append(row2)

    #print(data) # 參加抽獎人員
    #wait = input("Press Enter to continue.")
    # Randomly shuffle the data
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
    csv_writer.writerow(['Sequence','EmployeeID', 'Name', 'Gender','Gift'])
    for i, emp_data in enumerate(lucky_draw_result,start=1):
        #for member in number_of_times:
        #print(i,emp_data)
        csv_writer.writerow([i, emp_data[0],emp_data[1],emp_data[2],emp_data[4]])

    # Save the CSV to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as csv_temp_file:
        csv_temp_file_path = csv_temp_file.name
        #csv_temp_file.write(csv_output.getvalue().encode('utf-8'))
        csv_temp_file.write(csv_output.getvalue().encode('big5'))

       

    return render_template('result.html', teams=lucky_draw_result, csv_file_path=csv_temp_file_path, pdf_file_path=None)

@app.route('/download_csv')
def download_csv():
    csv_file_path = request.args.get('csv_file_path')
    return send_file(csv_file_path, as_attachment=True)

# @app.route('/download_pdf')
# def download_pdf():
#     pdf_file_path = request.args.get('pdf_file_path')
#     return send_file(pdf_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
