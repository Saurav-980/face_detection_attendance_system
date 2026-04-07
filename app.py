from flask import Flask, render_template, request, redirect, url_for
import cv2
import pickle
import numpy as np
import os
import csv
import time
import pandas as pd # Naya import
from datetime import datetime
from win32com.client import Dispatch
from sklearn.neighbors import KNeighborsClassifier

app = Flask(__name__)

# Dummy login credentials
VALID_EMAIL = "sauravshubham444@gmail.com"
VALID_PASSWORD = "04082003"

# Folders check karna
if not os.path.exists('data/'):
    os.makedirs('data/')
if not os.path.exists('Attendance/'):
    os.makedirs('Attendance/')

def speak(str1):
    try:
        speak_engine = Dispatch("SAPI.SpVoice")
        speak_engine.Speak(str1)
    except Exception as e:
        print(f"Speak Error: {e}")

@app.route('/')
def home():
    return render_template('Home.html')

@app.route('/login', methods=['POST'])
def login():
    email = request.form.get('email')
    password = request.form.get('password')
    if email == VALID_EMAIL and password == VALID_PASSWORD:
        return redirect(url_for('add_student'))
    else:
        return "<h1>Invalid Credentials!</h1><a href='/'>Try Again</a>"

# 1. ADD STUDENT ROUTE
@app.route('/add-student', methods=['GET', 'POST'])
def add_student():
    if request.method == 'POST':
        name = request.form.get('name')
        roll = request.form.get('roll')
        branch = request.form.get('branch')
        course = request.form.get('course')
        user_info = f"{name}|{roll}|{branch}|{course}"

        video = cv2.VideoCapture(0)
        facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')
        faces_data = []
        i = 0

        while True:
            ret, frame = video.read()
            if not ret: break
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = facedetect.detectMultiScale(gray, 1.3, 5)
            for (x, y, w, h) in faces:
                crop_img = frame[y:y+h, x:x+w, :]
                resized_img = cv2.resize(crop_img, (50, 50))
                if len(faces_data) < 100 and i % 10 == 0:
                    faces_data.append(resized_img)
                i += 1
                cv2.putText(frame, str(len(faces_data)), (50, 50), cv2.FONT_HERSHEY_COMPLEX, 1, (50, 50, 255), 1)
                cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 1)
            cv2.imshow("Registration Frame", frame)
            k = cv2.waitKey(1)
            if k == ord('q') or len(faces_data) == 100: break

        video.release()
        cv2.destroyAllWindows()

        faces_data = np.asarray(faces_data).reshape(100, -1)

        # Save Names
        names_path = 'data/names.pkl'
        if not os.path.isfile(names_path):
            names = [user_info] * 100
        else:
            with open(names_path, 'rb') as f: names = pickle.load(f)
            names = names + [user_info] * 100
        with open(names_path, 'wb') as f: pickle.dump(names, f)

        # Save Face Data
        faces_path = 'data/faces_data.pkl'
        if not os.path.isfile(faces_path):
            faces = faces_data
        else:
            with open(faces_path, 'rb') as f: faces = pickle.load(f)
            faces = np.append(faces, faces_data, axis=0)
        with open(faces_path, 'wb') as f: pickle.dump(faces, f)

        return f"<h1>Registration Successful for {name}!</h1><a href='/add-student'>Add Another</a> | <a href='/'>Go Home</a>"

    return render_template('add_stu.html')

# 2. MAKE ATTENDANCE ROUTE
@app.route('/make-attendance')
def make_attendance():
    if not os.path.exists('data/names.pkl') or not os.path.exists('data/faces_data.pkl'):
        return "<h1>Error: Dataset not found. Please register students first!</h1>"

    with open('data/names.pkl', 'rb') as f: LABELS = pickle.load(f)
    with open('data/faces_data.pkl', 'rb') as f: FACES = pickle.load(f)

    knn = KNeighborsClassifier(n_neighbors=5)
    knn.fit(FACES, LABELS)

    video = cv2.VideoCapture(0)
    facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')
    column_names = ['NAME', 'ROLL_NO', 'BRANCH', 'COURSE', 'TIME']

    while True:
        ret, frame = video.read()
        if not ret: break
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = facedetect.detectMultiScale(gray, 1.3, 5)
        current_attendance = None 

        for (x, y, w, h) in faces:
            crop_img = frame[y:y+h, x:x+w, :]
            resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
            output = knn.predict(resized_img)
            
            info = str(output[0]).split('|')
            name, roll, branch, course = info[0], info[1], info[2], info[3]
            
            ts = time.time()
            date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
            timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
            
            cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 0, 255), 1)
            cv2.putText(frame, name, (x, y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
            current_attendance = [name, roll, branch, course, timestamp]

        cv2.imshow("Attendance System - Press 'o' to Mark, 'q' to Exit", frame)
        k = cv2.waitKey(1)
        
        if k == ord('o') and current_attendance:
            speak("Attendance Taken")
            file_path = f"Attendance/Attendance_{date}.csv"
            exist = os.path.isfile(file_path)
            with open(file_path, "a", newline="") as csvfile:
                writer = csv.writer(csvfile)
                if not exist: writer.writerow(column_names)
                writer.writerow(current_attendance)
        
        if k == ord('q'): break

    video.release()
    cv2.destroyAllWindows()
    return redirect(url_for('home'))

# 3. REPORT VIEW ROUTE (Naya Dashboard Logic)
@app.route('/report')
def report():
    ts = time.time()
    date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
    file_path = f"Attendance/Attendance_{date}.csv"
    
    attendance_data = []
    if os.path.exists(file_path):
        try:
            df = pd.read_csv(file_path)
            # CSV data ko dashboard ke liye dictionary mein badalna
            attendance_data = df.to_dict(orient='records')
        except Exception as e:
            print(f"CSV Read Error: {e}")
            
    return render_template('report.html', data=attendance_data, date=date)

if __name__ == '__main__':
    app.run(debug=True)