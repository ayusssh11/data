from flask import Flask, render_template, request, redirect, url_for, send_file
from openpyxl import Workbook
from io import BytesIO
import os

app = Flask(__name__)
data = []

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form['name'].strip()
    email = request.form['email'].strip()
    address = request.form['address'].strip()

    if name and email and address:
        data.append({'Name': name, 'Email': email, 'Address': address})
        return redirect(url_for('index'))
    else:
        return "Fill all fields!", 400

@app.route('/download')
def download():
    # Create an in-memory Excel file
    output = BytesIO()
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Name", "Email", "Address"])  # Header row

    for row in data:
        sheet.append([row['Name'], row['Email'], row['Address']])

    workbook.save(output)
    output.seek(0)

    return send_file(output, download_name="user_data.xlsx", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)