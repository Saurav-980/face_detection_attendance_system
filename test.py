from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

def speak(str1):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str1)

video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

with open('data/names.pkl', 'rb') as f:
    LABELS = pickle.load(f)
with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

# Column names ab expand ho gaye hain
column_names = ['NAME', 'ROLL_NO', 'BRANCH', 'COURSE', 'TIME']

if not os.path.exists("Attendance"):
    os.makedirs("Attendance")

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)

    attendance = None 

    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        
        output = knn.predict(resized_img)
        
        # String ko split karke individual data nikalna
        info = str(output[0]).split('|')
        name, roll, branch, course = info[0], info[1], info[2], info[3]
        
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
        
        cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y-40), (x+w, y), (50, 50, 255), -1)
        cv2.putText(frame, name, (x, y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        
        attendance = [name, roll, branch, course, timestamp]

    cv2.imshow("Attendance System", frame)
    k = cv2.waitKey(1)
    
    if k == ord('o'):
        if attendance is not None:
            speak("Attendance Taken")
            file_path = "Attendance/Attendance_" + date + ".csv"
            exist = os.path.isfile(file_path)
            
            with open(file_path, "a", newline="") as csvfile: # Standard 'a' mode
                writer = csv.writer(csvfile)
                if not exist:
                    writer.writerow(column_names)
                writer.writerow(attendance)
            print(f"Attendance Recorded for {attendance[0]}")
        else:
            print("No face detected!")
            
    if k == ord('q'):
        break

video.release()
cv2.destroyAllWindows()