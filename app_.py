import csv
import random
from flask import Flask, render_template, request, make_response, send_file
from io import StringIO, BytesIO
from fpdf import FPDF
import tempfile

app = Flask(__name__, static_url_path='/static')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/randomize', methods=['POST'])
def randomize():
    csv_file = request.files['csv_file']
    num_teams = int(request.form['num_teams'])
    force_female_balance = 'force_female_balance' in request.form

    # Read the CSV file and parse the data
    data = []
    reader = csv.DictReader(StringIO(csv_file.read().decode('utf-8')))
    for row in reader:
        data.append(row)

    # Randomly shuffle the data
    random.shuffle(data)

    # Divide members into teams
    teams = [[] for _ in range(num_teams)]

    if force_female_balance:
        # Group female members first to ensure gender balance
        females = [member for member in data if member['Gender'] == '女']
        males = [member for member in data if member['Gender'] == '男']

        while females:
            for team in teams:
                if not females:
                    break
                team.append(females.pop())

        while males:
            for team in teams:
                if not males:
                    break
                team.append(males.pop())
    else:
        # Randomly assign members to teams
        for i, member in enumerate(data):
            team_index = i % num_teams
            teams[team_index].append(member)

    # Generate the result as a CSV file
    csv_output = StringIO()
    csv_writer = csv.writer(csv_output)
    csv_writer.writerow(['Team', 'Name', 'Gender'])
    for i, team in enumerate(teams, start=1):
        for member in team:
            csv_writer.writerow([i, member['Name'], member['Gender']])

    # Generate the result as a PDF
    pdf_output = BytesIO()
    pdf = FPDF()
    pdf.add_page()

    # Register the font
    pdf.add_font('simsun', '', './SimSun.ttf', uni=True)

    # Set the font for the PDF
    pdf.set_font('simsun', '', 12)

    # Add the content to the PDF
    pdf.cell(0, 10, '隨機分組結果', ln=True, align='C')
    pdf.ln(10)
    for i, team in enumerate(teams, start=1):
        pdf.cell(0, 10, f'團隊-{i}', ln=True)
        for member in team:
            pdf.cell(0, 8, f'{member["Name"]} ({member["Gender"]})', ln=True)
        pdf.ln(5)

    # Save the CSV to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as csv_temp_file:
        csv_temp_file_path = csv_temp_file.name
        csv_temp_file.write(csv_output.getvalue().encode('utf-8'))

    # Save the PDF to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as pdf_temp_file:
        pdf_temp_file_path = pdf_temp_file.name
        pdf_temp_file.write(pdf_output.getvalue())

    return render_template('result.html', teams=teams, csv_file_path=csv_temp_file_path, pdf_file_path=pdf_temp_file_path)

@app.route('/download_csv')
def download_csv():
    csv_file_path = request.args.get('csv_file_path')
    return send_file(csv_file_path, as_attachment=True)

@app.route('/download_pdf')
def download_pdf():
    pdf_file_path = request.args.get('pdf_file_path')
    return send_file(pdf_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)   
