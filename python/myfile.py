import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime

# Функция поиска расписания
def find_schedule_by_teacher_name(teacher_name, file_path, sheet_names):
    today = datetime.today().strftime('%Y-%m-%d')
    schedule = {}
    for sheet in sheet_names:
        try:
            data = pd.read_excel(file_path, sheet_name=sheet, header=None)
        # Остальная часть обработки листа...
        except ValueError:
            print(f"Лист '{sheet}' не найден в файле.")
            continue

        data = pd.read_excel(file_path, sheet_name=sheet, header=None)
        for index, row in data.iterrows():
            if teacher_name.lower() in str(row[3]).lower():
                tName = data.iloc[index][3]
                fIndex = 1
                while pd.isnull(data.iloc[index - fIndex][3]):
                    fIndex += 1
                subject = data.iloc[index - fIndex][3]
                subject_indx = index - fIndex
                    
                fIndex = 0
                while pd.isnull(data.iloc[index - fIndex][0]):
                    fIndex += 1

                day = data.iloc[index - fIndex][0]
                date = data.iloc[index - fIndex][1]

                time = []
                for time_index in range(subject_indx, index):
                    time_entry = data.iloc[time_index][2]
                    if pd.notnull(time_entry):
                        time.append(time_entry)

                if pd.notnull(date) and date != 'nan':
                    date = date.strftime('%Y-%m-%d')

                if date == today:
                    schedule_entry = f"{day}, {date}: {subject} в {time}, {tName}"
                    if sheet in schedule:
                        schedule[sheet].append(schedule_entry)
                    else:
                        schedule[sheet] = [schedule_entry]
    return schedule

# Функция для вызова диалога выбора файла
def open_file_dialog():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(0, file_path)

# Функция для обработки данных и вывода результатов
def process_and_show_schedule():
    teacher_name = teacher_name_entry.get()
    file_path = file_path_entry.get()
    if not file_path:
        messagebox.showerror("Ошибка", "Пожалуйста, выберите файл Excel")
        return
    if not teacher_name:
        messagebox.showerror("Ошибка", "Пожалуйста, введите имя преподавателя")
        return

    sheet_names = ['1 курс', '2 курс', '3 курс']  # Список листов для поиска
    schedule = find_schedule_by_teacher_name(teacher_name, file_path, sheet_names)
    output_text.delete('1.0', tk.END)
    if schedule:
        for course, schedule_list in schedule.items():
            output_text.insert(tk.END, f"Расписание для {course}:\n")
            for schedule_item in schedule_list:
                output_text.insert(tk.END, schedule_item + "\n")
            output_text.insert(tk.END, "\n")
    else:
        output_text.insert(tk.END, "На сегодня пар нет\n")

# Создание главного окна приложения
root = tk.Tk()
root.title("Расписание преподавателей")

# Элементы интерфейса
label = tk.Label(root, text="Введите фамилию преподавателя, формат Иванов И.И.:")
label.pack()

teacher_name_entry = tk.Entry(root)
teacher_name_entry.pack()

file_path_button = tk.Button(root, text="Выберите файл Excel", command=open_file_dialog)
file_path_button.pack()

file_path_entry = tk.Entry(root)
file_path_entry.pack()

process_button = tk.Button(root, text="Показать расписание", command=process_and_show_schedule)
process_button.pack()

output_text = tk.Text(root, height=10, width=50)
output_text.pack()

# Запуск основного цикла
root.mainloop()
