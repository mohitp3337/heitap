import xlrd
import os
from flask import Flask
from flask import render_template
from flask import request
from flask import url_for, redirect, flash
from pandas import *


app = Flask(__name__)

@app.route('/')
def home():
    return render_template('login.html')

@app.route('/userlogin', methods=['POST', 'GET'])
def userlogin():
    if request.method == 'POST':
        formData = request.form
        user = request.form.getlist('name')
        user_name = user[0]
        if request.form['name'] == 'admin' and request.form['password'] == 'admin':
            print("User successfully logged in")
        else:
            print("Invalid credentials")
            return redirect(url_for('home'))
    return render_template('result.html', user_name=user_name)

@app.route('/administration', methods=['POST', 'GET'])
def administration():
    return "Administration section"

@app.route('/testrun', methods=['POST', 'GET'])
def testrun():
    return "Test-run section"

@app.route('/testresult', methods=['POST', 'GET'])
def testresult():
    return "Test-result section"

@app.route('/logout', methods=['GET'])
def logout():
    #if request.method == 'POST':
    print("User logout successfully >>>>>>>>>>>")
    #flash('You were successfully logged out')
    return redirect(url_for('home'))

def read_xl_file():
    directory_path = os.path.dirname(__file__)
    read_file = os.path.join(directory_path, 'testcase.Mendix_Automation.xlsx')
    wb = xlrd.open_workbook(read_file)
    return wb

def get_excel_rows():
    wb = read_xl_file()
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    total_rows = sheet.nrows
    return total_rows

def get_xl_column_count():
    wb = read_xl_file()
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    column_counts = sheet.ncols
    return column_counts

def get_xl_column_names():
    wb = read_xl_file()
    sheet = wb.sheet_by_index(0)
    for column_name in range(sheet.ncols):
        #print(sheet.cell_value(0, column_name))
        column_names = sheet.cell_value(0, column_name)
        #print(column_names)
        return column_names

def get_tc_id_list():
    tc_id_list = []
    wb = read_xl_file()
    sheet = wb.sheet_by_index(0)
    for tc_id in range(sheet.nrows):
        tc_ids = sheet.cell_value(tc_id, 0)
        tc_id_list.append(tc_ids)
    return tc_id_list

def get_tc_title():
    tc_title= []
    wb = read_xl_file()
    sheet = wb.sheet_by_index(0)
    for title in range(sheet.nrows):
        tc_titles = sheet.cell_value(title, 1)
        tc_title.append(tc_titles)
    tc_title_list = tc_title[1:]
    return tc_title_list

def get_tc_title():
    tc_title= []
    wb = read_xl_file()
    sheet = wb.sheet_by_index(0)
    for title in range(sheet.nrows):
        tc_titles = sheet.cell_value(title, 1)
        tc_title.append(tc_titles)
    tc_title_list = tc_title[1:]
    return tc_title_list

@app.route('/testcase', methods=['POST', 'GET'])
def testcase():
    tc_list = get_tc_id_list()
    total_test_case = len(tc_list)
    tc_title = get_tc_title()
    #return "Test-case section"
    return render_template('testcase.html', tc_list=tc_list, tc_title=tc_title, total_test_case=total_test_case)

def excel_to_dict():
    directory_path = os.path.dirname(__file__)
    read_file = os.path.join(directory_path, 'testcase.Mendix_Automation.xlsx')
    xls = ExcelFile(read_file)
    data = xls.parse(xls.sheet_names[0])

@app.route('/result', methods=['POST', 'GET'])
def mohit():
    #return "mohit>>>>>>>>>>>"
    redirect(url_for('home'))



#get_xl_column_count()
#get_xl_column_names()
#get_tc_id_list()
get_tc_title()

if __name__ == '__main__':
    app.run()
