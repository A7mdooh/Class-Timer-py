import tkinter as tk
from tkinter import ttk
from datetime import datetime
import pandas as pd
import time
import pygame
import os
import random
import cv2
from threading import Thread
from PIL import Image, ImageTk
import subprocess

# تهيئة مكتبة pygame
pygame.mixer.init()

# إنشاء نافذة Tkinter
root = tk.Tk()

root.title("عرض الأحداث حسب اليوم")

# إنشاء إطار لوضع الشعار وساعة المدرسة والساعة الرقمية
header_frame = tk.Frame(root)
header_frame.pack(side="top", fill="x")

# إضافة شعار (logo) وتصغير حجمه ووضعه في الجزء الأيمن من الشريط
logo_image = tk.PhotoImage(file="logo.png")
logo_image = logo_image.subsample(2, 2)  # تصغير حجم الصورة بنسبة 1/2
logo_label = tk.Label(header_frame, image=logo_image)
logo_label.pack(side="right")

# تحديث مظهر عنصر Label الذي يعرض اسم المدرسة
try:
    with open("info.txt", "r", encoding="utf-8") as info_file:
        manager_data = info_file.readlines()
        school_manager_name = manager_data[0].strip()
        assistant_manager_name = manager_data[1].strip()
        school_name = manager_data[2].strip()
except FileNotFoundError:
    school_manager_name = "اسم المدير"
    assistant_manager_name = "اسم المدير المساعد"
    school_name = "اسم المدرسة"

school_name_label = tk.Label(header_frame, text=school_name, font=("Cairo", 24, "bold"),
                             fg="#006400")  # تحديث الخصائص هنا
school_name_label.pack(side="top", pady=10)

# إضافة عنوان ساعة المدرسة وتعيين موقعه في الجزء الأوسط من الشريط
title_label = tk.Label(header_frame, text="ساعة المدرسة", font=("Cairo", 18), fg="red")
title_label.pack(side="top", pady=10)

# إضافة الساعة الرقمية وتعيين موقعها في الجزء الأيسر من الشريط
time_label = tk.Label(header_frame, font=("Cairo", 30), fg="blue")
time_label.pack(side="left", padx=20, pady=10)

# إنشاء إطار لوضع بيانات المدير والمدير المساعد وزر إنهاء التطبيق في الجزء السفلي المنتصف من الشاشة
managers_and_exit_frame = tk.Frame(root)
managers_and_exit_frame.pack(side="bottom", fill="x", pady=10)

# إضافة مدير المدرسة إلى الجهة اليمنى المنتصف
school_manager_label = tk.Label(managers_and_exit_frame, text=f"مدير المدرسة: {school_manager_name}",
                                font=("Cairo", 12, "bold"), fg="#006400")
school_manager_label.pack(side="right", padx=20)

# إضافة مدير المدرسة المساعد إلى الجهة اليسرى المنتصف
assistant_manager_label = tk.Label(managers_and_exit_frame, text=f"المدير  المساعد: {assistant_manager_name}",
                                   font=("Cairo", 12, "bold"), fg="#006400")
assistant_manager_label.pack(side="left", padx=20)

# إضافة زر إنهاء  في الوسط
exit_button = tk.Button(managers_and_exit_frame, text="إنهـــاء ", command=root.quit, font=("Cairo", 12), bg="red",
                        fg="black")
exit_button.pack(fill="x", pady=10)

# إنشاء إطار لوضع اسم المصمم في الجزء الأسفل وسط الشاشة
designer_frame = tk.Frame(root)
designer_frame.pack(side="bottom", fill="x")

# إضافة مكون Label جديد لعرض اسم المصمم وتعيين موقعه في وسط الشاشة
designer_name_label = tk.Label(designer_frame, text="برمجة وتصميم : أ / أحمد بن عبدالله الضامري  ",
                               font=("Cairo", 10, "bold"), fg="red")
designer_name_label.pack()

# تعيين حجم النافذة إلى كامل الشاشة
root.attributes('-fullscreen', True)

# إنشاء واجهة المستخدم
day_label = tk.Label(root, text="اختر يومًا من الأسبوع أو معرض:")
day_label.pack(pady=0)

days = ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "معرض الصور", "معرض الفيديو", "إيقاف"]
day_var = tk.StringVar()
day_dropdown = ttk.Combobox(root, textvariable=day_var, values=days)
day_dropdown.pack(pady=10)
day_dropdown.set(days[0])

# إنشاء Treeview لعرض الأحداث
event_frame = tk.Frame(root)
event_frame.pack(padx=20, pady=0, fill="both", expand=True)

# إنشاء Treeview لعرض الأحداث بدون شريط التمرير
event_tree = ttk.Treeview(event_frame, columns=("event_name", "start_time", "end_time", "teacher", "class_name"),
                          height=24)
style = ttk.Style()
style.theme_use("clam")  # استخدام ثيم يدعم تغيير الألوان بشكل أفضل

event_tree["show"] = "headings"
event_tree.heading("event_name", text="الحدث")
event_tree.heading("start_time", text="وقت البداية")
event_tree.heading("end_time", text="وقت النهاية")
event_tree.heading("teacher", text="المعلم")
event_tree.heading("class_name", text="الصف")
event_tree.pack(padx=20, pady=5, fill="both", expand=True)

# نافذة جديدة لعرض الصور بكامل الشاشة
photo_window = None
current_photo_index = 0
paused = False

# عرض معرض الصور

def show_photo_gallery():
    global photo_window, current_photo_index, paused
    photo_folder = "Photo gallery"
    if not os.path.exists(photo_folder):
        return

    photos = [f for f in os.listdir(photo_folder) if f.endswith(('.png', '.jpg', '.jpeg'))]
    if not photos:
        return

    current_photo_index = 0
    paused = False

    # إنشاء نافذة جديدة لعرض الصور
    if photo_window is None or not photo_window.winfo_exists():
        photo_window = tk.Toplevel(root)
        photo_window.attributes('-fullscreen', True)
        photo_window.title("معرض الصور")

        # إضافة زر إنهاء في نافذة الصور
        exit_button = tk.Button(photo_window, text="إنهاء العرض", font=("Cairo", 12), command=exit_photo_gallery)
        exit_button.pack(side="top", pady=10)

        # إضافة إطار لعرض الصورة
        photo_label = tk.Label(photo_window)
        photo_label.pack(fill="both", expand=True)

        # عرض الصور بشكل متسلسل
        def display_photo(photo_index):
            if paused or not photo_window.winfo_exists():
                return

            photo_file = os.path.join(photo_folder, photos[photo_index])
            image = Image.open(photo_file)
            image = image.resize((photo_window.winfo_width(), photo_window.winfo_height()), Image.LANCZOS)
            photo = ImageTk.PhotoImage(image)

            photo_label.config(image=photo)
            photo_label.image = photo
            photo_window.after(5000, lambda: display_photo((photo_index + 1) % len(photos)))

        display_photo(current_photo_index)

# إنهاء معرض الصور والعودة إلى الشاشة الرئيسية

def exit_photo_gallery():
    global photo_window
    if photo_window is not None and photo_window.winfo_exists():
        photo_window.destroy()
        photo_window = None
    event_frame.pack(padx=20, pady=0, fill="both", expand=True)
    day_var.set("إيقاف")  # تعيين القيمة إلى "إيقاف" لإيقاف الشجرة

# نافذة جديدة لعرض الفيديوهات بكامل الشاشة
video_window = None
current_video_index = 0
video_process = None

# تعديل النهج المستخدم لعرض الفيديوهات باستخدام VLC لضمان تشغيل الصوت والفيديو بشكل كامل

def show_video_gallery():
    global video_window, current_video_index, video_process
    video_folder = "Video gallery"
    if not os.path.exists(video_folder):
        print("لا يوجد مجلد الفيديوهات.")
        return

    videos = [f for f in os.listdir(video_folder) if f.endswith('.mp4')]
    if not videos:
        print("لا توجد فيديوهات بصيغة .mp4.")
        return

    current_video_index = 0

    # إنشاء نافذة جديدة لعرض الفيديو
    if video_window is None or not video_window.winfo_exists():
        video_window = tk.Toplevel(root)
        video_window.attributes('-fullscreen', True)
        video_window.title("معرض الفيديو")

        # إضافة زر إنهاء في نافذة الفيديو
        exit_button = tk.Button(video_window, text="إنهاء العرض", font=("Cairo", 12), command=exit_video_gallery)
        exit_button.pack(side="top", pady=10)

        # تشغيل الفيديو بشكل متسلسل باستخدام VLC لضمان تشغيل الصوت والفيديو معًا
        def play_video(video_index):
            global video_process
            if video_window is None or not video_window.winfo_exists():
                return

            video_file = os.path.join(video_folder, videos[video_index])
            print(f"تشغيل الفيديو: {video_file}")

            try:
                # تشغيل الفيديو باستخدام VLC
                video_process = subprocess.Popen(
                    [r"C:\Program Files\VideoLAN\VLC\vlc.exe", "--fullscreen", "--play-and-exit", video_file]
                )
                # الانتقال إلى الفيديو التالي عند انتهاء الفيديو الحالي
                video_duration = cv2.VideoCapture(video_file).get(cv2.CAP_PROP_FRAME_COUNT) / cv2.VideoCapture(video_file).get(cv2.CAP_PROP_FPS)
                video_window.after(int(video_duration * 1000), lambda: play_video((video_index + 1) % len(videos)))
            except FileNotFoundError:
                print("لم يتم العثور على VLC لتشغيل الفيديو. يرجى التحقق من تثبيته.")

        play_video(current_video_index)

# إنهاء معرض الفيديو والعودة إلى الشاشة الرئيسية
def exit_video_gallery():
    global video_window, video_process
    if video_window is not None and video_window.winfo_exists():
        video_window.destroy()
        video_window = None
    if video_process is not None:
        video_process.terminate()
        video_process = None
    event_frame.pack(padx=20, pady=0, fill="both", expand=True)
    day_var.set("إيقاف")  # تعيين القيمة إلى "إيقاف" لإيقاف الشجرة

# تحديث مظهر عنصر Label الذي يعرض اسم المدرسة
def update_school_name_label():
    try:
        with open("info.txt", "r", encoding="utf-8") as info_file:
            manager_data = info_file.readlines()
            school_name = manager_data[2].strip()
            school_name_label.config(text=school_name)
    except FileNotFoundError:
        pass

# قراءة ملف Excel وملء جدول الأحداث
def load_events():
    try:
        selected_day = day_var.get()
        if selected_day == "معرض الصور":
            show_photo_gallery()
            return
        elif selected_day == "معرض الفيديو":
            show_video_gallery()
            return
        elif selected_day == "إيقاف":
            event_tree.delete(*event_tree.get_children())
            return

        # إظهار الشجرة وإخفاء معرض الصور
        event_frame.pack(padx=20, pady=0, fill="both", expand=True)

        df = pd.read_excel("events.xlsx")
        current_time = datetime.now().strftime("%H:%M:%S")

        matching_events = []

        for index, row in df.iterrows():
            start_time_str = row["start_time"].strftime("%H:%M:%S")
            end_time_str = row["end_time"].strftime("%H:%M:%S")

            if row["day"] == selected_day and start_time_str <= current_time <= end_time_str:
                matching_events.append(
                    (row["event_name"], start_time_str, end_time_str, row["teacher"], row["class_name"]))

        event_tree.delete(*event_tree.get_children())

        if matching_events:
            # قائمة بألوان مميزة للأحداث
            event_colors = ["#FF5733", "#33FF57", "#5733FF", "#FFFF33", "#33FFFF", "#ccffcc", "#3366FF", "#cccccc",
                            "#99cccc", "#66ffcc", "#669966", "#ffff00", "#99ffff", "#00ccff", "#ccffcc", "#ffcc99",
                            "#ffcccc", "#ccff99", "#99ffcc", "#ccffff"]

            for idx, event in enumerate(matching_events):
                event_id = event_tree.insert("", "end", values=event)
                event_color = event_colors[idx % len(event_colors)]  # اختيار لون من القائمة بناءً على مؤشر الحدث
                event_tree.tag_configure(f"event_{idx}", background=event_color,
                                         foreground="black")  # قم بتعريف الوسم وتحديد لون الخلفية
                event_tree.item(event_id, tags=(f"event_{idx}",))  # قم بتطبيق الوسم على الحدث
                play_start_sound()
                event_duration = (datetime.strptime(event[2], "%H:%M:%S") - datetime.strptime(event[1],
                                                                                              "%H:%M:%S")).seconds * 1000
                root.after(event_duration, play_end_sound)
                close_video_if_playing()  # إغلاق الفيديو عند بدء حدث جديد
        else:
            event_tree.insert("", "end", values=("لا توجد أحداث بالوقت الحالي", "", "", "", ""))
            play_video()  # تشغيل الفيديو عند عدم وجود أحداث

    except Exception as e:
        event_tree.insert("", "end", values=("خطأ", str(e), "", "", ""))

video_process = None
video_playing = False

# تشغيل الفيديو عند عدم وجود أحدث
def play_video():
    global video_playing, video_process
    if video_playing:
        return

    video_folder = "Short video"
    if not os.path.exists(video_folder):
        return

    videos = [f for f in os.listdir(video_folder) if f.endswith(".mp4")]
    if not videos:
        return

    # اختيار فيديو عشوائي لتشغيله
    video_file = os.path.join(video_folder, random.choice(videos))

    try:
        # تشغيل الفيديو باستخدام مشغل خارجي (مثل VLC) لضمان تشغيل الصوت
        video_process = subprocess.Popen(
            [r"C:\Program Files\VideoLAN\VLC\vlc.exe", "--fullscreen", "--play-and-exit", video_file]
        )
        video_playing = True
    except FileNotFoundError:
        event_tree.insert("", "end",
                          values=("خطأ", "لم يتم العثور على VLC لتشغيل الفيديو. يرجى التحقق من تثبيته.", "", "", ""))

# إغلاق الفيديو إذا كان يعمل
def close_video_if_playing():
    global video_playing, video_process
    if video_playing:
        if video_process is not None:
            video_process.terminate()
            video_process = None
        video_playing = False

# تحديث الساعة الرقمية بشكل دوري وتحميل الأحداث
def update_time_and_load_events():
    current_time = time.strftime("%H:%M:%S")  # استخدام مكتبة time للتحسين في عرض الوقت
    time_label.config(text=current_time)
    load_events()
    update_school_name_label()  # تحديث اسم المدرسة

    root.after(1000, update_time_and_load_events)

# تشغيل الملف الصوتي لبداية الحدث
start_sound_played = False

def play_start_sound():
    global start_sound_played
    if not start_sound_played:
        try:
            pygame.mixer.music.load("start_sound.mp3")  # استبدل "start_sound.mp3" بملف الصوت الخاص ببداية الحدث
            pygame.mixer.music.play()
            start_sound_played = True
        except Exception as e:
            print("خطأ في تشغيل ملف الصوت:", str(e))

# تشغيل الملف الصوتي لنهاية الحدث مرة واحدة فقط
def play_end_sound():
    global start_sound_played
    try:
        if start_sound_played:
            pygame.mixer.music.load("end_sound.mp3")  # استبدل "end_sound.mp3" بملف الصوت الخاص بنهاية الحدث
            pygame.mixer.music.play()
            start_sound_played = False
    except Exception as e:
        print("خطأ في تشغيل ملف الصوت:", str(e))

# بدء تشغيل تحديث الوقت وتحميل الأحداث
update_time_and_load_events()

# تشغيل التطبيق
root.mainloop()
