import os
import base64
import pandas as pd
import requests
import tkinter as tk
from tkinter import filedialog, messagebox

def send_emails(folder_path, excel_path, sender_email, token, email_text):
    try:
        # Загрузка данных из Excel
        data = pd.read_excel(excel_path)

        if not {'ФИО', 'Email'}.issubset(data.columns):
            raise ValueError("Excel файл должен содержать колонки 'ФИО' и 'Email'")

        base_url = "https://graph.microsoft.com/v1.0/me/sendMail"

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        for _, row in data.iterrows():
            full_name = row['ФИО']
            recipient_email = row['Email']

            # Поиск файла
            pdf_file_name = f"{full_name}.pdf"
            pdf_file_path = os.path.join(folder_path, pdf_file_name)

            if not os.path.exists(pdf_file_path):
                print(f"Файл {pdf_file_name} не найден. Пропускаем...")
                continue

            with open(pdf_file_path, "rb") as f:
                pdf_content = f.read()

            encoded_pdf = base64.b64encode(pdf_content).decode('utf-8')

            email_body = {
                "message": {
                    "subject": "Ваш документ",
                    "body": {
                        "contentType": "Text",
                        "content": email_text.replace("{full_name}", full_name)  # Заменяем плейсхолдер на имя
                    },
                    "toRecipients": [
                        {"emailAddress": {"address": recipient_email}}
                    ],
                    "attachments": [
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "name": pdf_file_name,
                            "contentBytes": encoded_pdf
                        }
                    ]
                }
            }

            response = requests.post(base_url, headers=headers, json=email_body)

            if response.status_code == 202:
                print(f"Письмо для {full_name} ({recipient_email}) успешно отправлено.")
            else:
                print(f"Ошибка отправки для {full_name} ({recipient_email}): {response.status_code}, {response.text}")

        messagebox.showinfo("Успех", "Рассылка завершена!")

    except Exception as e:
        messagebox.showerror("Ошибка", str(e))


def browse_folder():
    folder_selected = filedialog.askdirectory()
    folder_path.set(folder_selected)

def browse_excel():
    file_selected = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    excel_path.set(file_selected)

def start_sending():
    folder = folder_path.get()
    excel = excel_path.get()
    email = sender_email.get()
    token = api_token.get()
    email_text = email_text_entry.get("1.0", tk.END).strip()  # Получаем текст из текстового поля

    if not folder or not excel or not email or not token or not email_text:
        messagebox.showwarning("Ошибка", "Все поля должны быть заполнены")
        return

    send_emails(folder, excel, email, token, email_text)

# Создание GUI
root = tk.Tk()
root.title("Приложение для рассылки")

folder_path = tk.StringVar()
excel_path = tk.StringVar()
sender_email = tk.StringVar()
api_token = tk.StringVar()

# Поля ввода и кнопки
tk.Label(root, text="Почта отправителя:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=sender_email, width=40).grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="API Токен:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=api_token, width=40, show="*").grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Папка с PDF файлами:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=folder_path, width=40).grid(row=2, column=1, padx=10, pady=5)
tk.Button(root, text="Обзор", command=browse_folder).grid(row=2, column=2, padx=10, pady=5)

tk.Label(root, text="Excel файл:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=excel_path, width=40).grid(row=3, column=1, padx=10, pady=5)
tk.Button(root, text="Обзор", command=browse_excel).grid(row=3, column=2, padx=10, pady=5)

tk.Label(root, text="Текст письма:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
email_text_entry = tk.Text(root, height=5, width=40)
email_text_entry.grid(row=4, column=1, padx=10, pady=5)

tk.Button(root, text="Начать рассылку", command=start_sending, bg="green", fg="white").grid(row=5, column=1, pady=20)

root.mainloop()