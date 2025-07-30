from flask import Flask, render_template, request, redirect, url_for
from pymongo import MongoClient
from bson.objectid import ObjectId
import pandas as pd

app = Flask(__name__)

# Connect to MongoDB
client = MongoClient("mongodb://localhost:27017/")
db = client["mydatabase"]
collection = db["mycollection"]

# Export to Excel function
def export_to_excel():
    entries = list(collection.find())
    df = pd.DataFrame(entries)
    if '_id' in df:
        df.drop(columns=['_id'], inplace=True)
    df.to_excel("student_records.xlsx", index=False)

# Home Page
@app.route('/')
def home():
    return render_template('home.html')

# Student View - read-only
@app.route('/student')
def student_view():
    entries = list(collection.find())
    return render_template('student.html', entries=entries)

# Teacher View - add/edit/delete
@app.route('/teacher', methods=['GET', 'POST'])
def teacher_view():
    if request.method == 'POST':
        student_name = request.form['student_name']
        student_class = request.form['student_class']
        subject = request.form['subject']
        roll_number = request.form['roll_number']
        marks = request.form['marks']

        if student_name and student_class and subject and roll_number and marks:
            data = {
                'student_name': student_name,
                'student_class': student_class,
                'subject': subject,
                'roll_number': roll_number,
                'marks': marks
            }
            collection.insert_one(data)
            export_to_excel()

        return redirect(url_for('teacher_view'))

    entries = list(collection.find())
    return render_template('teacher.html', entries=entries)

# Delete Student
@app.route('/delete/<id>', methods=['POST'])
def delete(id):
    collection.delete_one({'_id': ObjectId(id)})
    export_to_excel()
    return redirect(url_for('teacher_view'))

if __name__ == '__main__':
    app.run(debug=True)
