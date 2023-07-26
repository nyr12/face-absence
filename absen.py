import tkinter as tk
import cv2
import os
import openpyxl
from openpyxl import Workbook
from PIL import Image, ImageTk
import threading

# Fungsi untuk menambahkan data ke file Excel
def add_data_to_excel(name, student_id, class_name):
    excel_file = 'D:\\absensi\\data_absensi.xlsx'
    if not os.path.exists(excel_file):
        wb = Workbook()
        ws = wb.active
        ws.append(['Nama', 'Nomor Induk', 'Kelas'])
    else:
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
    ws.append([name, student_id, class_name])
    wb.save(excel_file)

# Fungsi untuk registrasi wajah
def register_face():
    name = entry_name.get()
    student_id = entry_id.get()
    class_name = entry_class.get()

    global capture_flag, captured_image
    capture_flag = True
    captured_image = []

    # Simpan data ke Excel
    add_data_to_excel(name, student_id, class_name)

    # Aktifkan kamera dan simpan gambar wajah
    def capture_and_display():
        camera = cv2.VideoCapture(0)
        while True:
            ret, frame = camera.read()
            if not ret:
                break

        # Konversi gambar dari OpenCV (BGR) ke format yang dapat ditampilkan di tkinter (RGB)
            rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            image = Image.fromarray(rgb_frame)
            imgtk = ImageTk.PhotoImage(image=image)

            # Tampilkan rekaman di GUI
            label_camera.config(image=imgtk)
            label_camera.img = imgtk

            if capture_flag:
                captured_image.append(frame)

            root.update()

        camera.release()

# Fungsi untuk mengenali wajah
def recognize_face():
    # Aktifkan kamera dan ambil gambar
    camera = cv2.VideoCapture(0)
    ret, frame = camera.read()
    camera.release()

    global capture_flag, captured_image
    capture_flag = True
    captured_image = []

    # Lakukan proses pengenalan wajah menggunakan face-recognition
    # Code pengenalan wajah akan ditambahkan di sini sesuai dokumentasi face-recognition.

    # Simpan hasil pengenalan wajah ke Excel
    # Code simpan hasil pengenalan akan ditambahkan di sini sesuai kebutuhan Anda.

# Buat GUI
root = tk.Tk()
root.title("Aplikasi Absensi")

label_name = tk.Label(root, text="Nama:")
label_name.pack()
entry_name = tk.Entry(root)
entry_name.pack()

label_id = tk.Label(root, text="Nomor Induk:")
label_id.pack()
entry_id = tk.Entry(root)
entry_id.pack()

label_class = tk.Label(root, text="Kelas:")
label_class.pack()
entry_class = tk.Entry(root)
entry_class.pack()

button_register = tk.Button(root, text="Registrasi Wajah", command=register_face)
button_register.pack()

button_recognize = tk.Button(root, text="Pengenalan Otomatis", command=recognize_face)
button_recognize.pack()

root.mainloop()
