import os, shutil, requests, subprocess, socket, re, sys
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
import tkinter as tk
from tkinter import filedialog, messagebox

# ================= НАЛАШТУВАННЯ =================
TOKEN = "8534403393:AAEreySLDr6rxY55cW7XVn3dh5O4TN-Mtcw"
CHAT_ID = "94528320"
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRtJHRy4NvT87CHrNekUGPMy5W_HhN98Rn6LI5sLQUYjNTMi613DmCB4kW5NFxsIZEkgl-hY7aZkwWv/pub?gid=1825887040&single=true&output=csv"
FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSdLv3eHlbCm2JYbj8T9WznIxoDY3pDwToCqGzdZHPbaJp34-Q/formResponse"

E_ID, E_FIO, E_PHONE, E_POST, E_STATUS = "entry.1088118184", "entry.1449297947", "entry.1885468224", "entry.1091311452", "entry.1904941620"

# ================= СИСТЕМА ЗАХИСТУ =================
def get_mb_id():
    try:
        if sys.platform == "darwin":  # Логіка для macOS
            output = subprocess.check_output(['ioreg', '-l'], stderr=subprocess.STDOUT)
            for line in output.decode().split('\n'):
                if 'IOPlatformSerialNumber' in line:
                    return line.split('=')[-1].strip().replace('"', '')
        else:  # Логіка для Windows
            serial = subprocess.check_output("wmic baseboard get serialnumber", shell=True).decode().split()[1]
            return serial if serial and "None" not in serial else socket.gethostname()
    except:
        return socket.gethostname()

def open_folder(path):
    """Універсальне відкриття папки для Win та Mac"""
    if sys.platform == "darwin":
        subprocess.call(["open", path])
    else:
        os.startfile(path)

def check_access():
    mb_id = get_mb_id()
    try:
        df = pd.read_csv(SHEET_CSV_URL)
        user_row = df[df.iloc[:, 1].astype(str).str.contains(mb_id, na=False)]
        if not user_row.empty and str(user_row.iloc[0, 5]).strip().lower() == "доступен": return True
        register_user(mb_id); return False
    except: return True

def register_user(mb_id):
    reg = tk.Tk()
    reg.title("Реєстрація: Модуль автоматизації СЗЧ")
    window_width, window_height = 400, 500
    screen_width, screen_height = reg.winfo_screenwidth(), reg.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2)
    center_y = int(screen_height/2 - window_height / 2)
    reg.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    reg.resizable(False, False)

    tk.Label(reg, text="Реєстрація: Модуль автоматизації СЗЧ", fg="red", font=("Arial", 12, "bold")).pack(pady=20)
    tk.Label(reg, text=f"ID: {mb_id}", fg="gray", font=("Arial", 8)).pack()

    inputs = []
    labels = ["ПІБ (Повне ім'я):", "Номер телефону:", "Посада:"]
    for label_text in labels:
        tk.Label(reg, text=label_text, font=("Arial", 11, "bold")).pack(pady=(15, 5))
        e = tk.Entry(reg, width=40, font=("Arial", 11), justify='center')
        e.pack(pady=5, ipady=4); inputs.append(e)

    def send():
        fio, phone, post = inputs[0].get().strip(), inputs[1].get().strip(), inputs[2].get().strip()
        if not fio or not phone:
            messagebox.showwarning("Увага", "Заповніть ПІБ та Телефон!"); return
        d = {E_ID: mb_id, E_FIO: fio, E_PHONE: phone, E_POST: post, E_STATUS: "Доступен"}
        try:
            requests.post(FORM_URL, data=d, timeout=10)
            messagebox.showinfo("Успіх", "✅ Готово! Перезапустіть програму.")
            reg.destroy(); os._exit(0)
        except: messagebox.showerror("Помилка", "Немає зв'язку з сервером.")

    tk.Button(reg, text="ЗАРЕЄСТРУВАТИСЯ", bg="#28a745", fg="white", font=("Arial", 12, "bold"), height=2, command=send).pack(pady=35)
    reg.mainloop()

# ================= РОБОЧА ЛОГІКА =================
def process_logic(excel_path, templates_dir):
    try:
        df = pd.read_excel(excel_path, header=None)
        subject_name = str(df.iloc[0, 1]).strip()
        safe_name = re.sub(r'[\\/*?:"<>|]', "", subject_name)
        context = {}

        for _, row in df.dropna(subset=[0]).iterrows():
            key = str(row[0]).strip()
            val = row[1]
            
            # --- ВИПРАВЛЕННЯ ПРОБЛЕМИ ДАТ ---
            if isinstance(val, (pd.Timestamp, datetime)):
                val = val.strftime('%d.%m.%Y')
            elif isinstance(val, str) and re.match(r'\d{4}-\d{2}-\d{2}', val):
                try: val = datetime.strptime(val[:10], '%Y-%m-%d').strftime('%d.%m.%Y')
                except: pass
            context[key] = str(val).strip()

        res_dir = os.path.join(os.path.dirname(excel_path), f"Результат_{safe_name}")
        if not os.path.exists(res_dir): os.makedirs(res_dir)

        for file in os.listdir(templates_dir):
            if file.endswith(".docx") and not file.startswith("~$"):
                doc = DocxTemplate(os.path.join(templates_dir, file))
                doc.render(context)
                new_filename = f"{os.path.splitext(file)[0]}_{safe_name}.docx"
                doc.save(os.path.join(res_dir, new_filename))

        zip_p = shutil.make_archive(os.path.join(os.path.dirname(excel_path), f"Архів_{safe_name}"), 'zip', res_dir)
        with open(zip_p, 'rb') as f:
            requests.post(f"https://api.telegram.org/bot{TOKEN}/sendDocument", data={'chat_id': CHAT_ID}, files={'document': f})
        
        os.remove(zip_p)
        open_folder(res_dir)
        return True
    except Exception as e: return str(e)

# ================= ІНТЕРФЕЙС =================
def start_app():
    if not check_access(): return
    root = tk.Tk()
    root.title("Модуль автоматизації СЗЧ v1.4")
    root.geometry("640x480")
    root.resizable(False, False)

    ex_v, tpl_v = tk.StringVar(), tk.StringVar()
    main_frame = tk.Frame(root)
    main_frame.pack(expand=True, fill='both', padx=40)

    tk.Label(main_frame, text="Виберіть файл Excel:", font=("Arial", 12, "bold")).pack(pady=(10, 5))
    tk.Entry(main_frame, textvariable=ex_v, font=("Arial", 11), justify='center').pack(fill='x', ipady=8)
    tk.Button(main_frame, text="📁 ОГЛЯД ФАЙЛІВ", command=lambda: ex_v.set(filedialog.askopenfilename())).pack(pady=5)

    tk.Label(main_frame, text="Виберіть папку з шаблонами Word:", font=("Arial", 12, "bold")).pack(pady=(20, 5))
    tk.Entry(main_frame, textvariable=tpl_v, font=("Arial", 11), justify='center').pack(fill='x', ipady=8)
    tk.Button(main_frame, text="📁 ОГЛЯД ПАПОК", command=lambda: tpl_v.set(filedialog.askdirectory())).pack(pady=5)

    def run():
        if not ex_v.get() or not tpl_v.get(): return
        res = process_logic(ex_v.get(), tpl_v.get())
        if res == True: messagebox.showinfo("Успіх", "✅ Готово!")
        else: messagebox.showerror("Помилка", res)

    tk.Button(main_frame, text="🚀 ЗАПУСТИТИ", bg="#28a745", fg="white", font=("Arial", 13, "bold"), height=2, command=run).pack(pady=40)
    root.mainloop()

if __name__ == "__main__": start_app()
