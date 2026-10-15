import tkinter as tk
from tkinter import ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg  # модуль для встраивания графиков в tkinter
import openpyxl

# Функция для создания диаграммы
def plot_pie_chart(graph_values, title, frame):
    # Извлекаем названия и количества из данных для графика
    labels = [item[0] for item in graph_values]
    sizes = [item[1] for item in graph_values]

    # Создаем фигуру и оси для графика
    fig, ax = plt.subplots(figsize=(8, 8))
    # Строим круговую диаграмму
    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140)
    # Устанавливаем заголовок графика
    ax.set_title(title)
    # Обеспечиваем круговую форму диаграммы
    ax.axis('equal')
    # Создаем canvas для отображения графика в tkinter
    canvas = FigureCanvasTkAgg(fig, master=frame)
    # Отображаем canvas
    canvas.draw()
    # Упаковываем canvas в frame
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)


# Функция для отображения выбранного графика
def show_graph(choice):
    global values, names  # Используем глобальные переменные с данными и названиями
    # Преобразуем в целое число
    choice = int(choice.split(':')[0]) - 1
    answers = set()  # Создаем множество для хранения уникальных ответов

    # Работа с числовыми значениями и строками без ';'
    if type(values[choice][0]) != str or ';' not in str(values[choice][0]):
        # Если нет разделителя, просто добавляем значения в множество
        for j in values[choice]:
            if j is not None: # В одной из строк есть None, поэтому делаем проверку
                answers.add(j)

        graph_values = []  # Создаем список для хранения данных для графика
        for k in answers:
            graph_values.append([k, values[choice].count(k)])  # Добавляем значение и количество его повторений
    else:
        # Если есть разделитель, разделяем строку на отдельные ответы
        for m in range(len(values[choice])):
            values[choice][m] = str(values[choice][m]).split(';')  # Разделяем строку по ';'

        res_values = []  # Создаем список для хранения всех разделенных ответов
        for l in values[choice]:
            for q in l:
                if q != '':
                    res_values.append(q)  # Добавляем каждый ответ в список

        for w in res_values:
            answers.add(w)  # Добавляем уникальные ответы в множество

        graph_values = []  # Создаем список для хранения данных для графика
        for z in answers:
            graph_values.append([z, res_values.count(z)])  # Добавляем значение и количество его повторений

    # Очищаем frame с графиком перед отображением нового
    for widget in graph_frame.winfo_children():
        widget.destroy()

    # Отображаем круговую диаграмму
    plot_pie_chart(graph_values, names[choice], graph_frame)


# Загрузка данных из Excel
workbook = openpyxl.load_workbook('/Users/larionovvladimir/Downloads/Межпоколенческие_особенности_доверия_цифровым_технологиям1_101.xlsx')
sheet = workbook.active

# Cписок названий для графиков
names = ["Знакомы ли Вы с понятием 'цифровая грамотность'?", "Как Вы оцениваете свой уровень цифровой грамотности (где '1' - очень низкий, а '5' - очень высокий)?", "Как Вы относитесь к современным цифровым технологиям по шкале от одного до пяти (где '1' - резко негативно, а '5' - очень положительно)?", "Как Вы оцениваете уровень кибербезопасности в современном интернет-пространстве (по шкале от одного до пяти, где '1' - очень низкий, а '5' - очень высокий", "Доверяете ли Вы свои конфиденциальные данные интернету?", "В каких аспектах Вы доверяете цифровым технологиям?", "Какие проблемы Вы считаете наиболее важными в использовании цифровых", "Как Вы считаете, будут ли цифровые технологии по-прежнему актуальны через достаточно большой промежуток времени (например, через 50 лет)?", "Согласны ли Вы со следующим утверждением: 'Цифровые технологии очень полезны, поэтому необходимо и дальше их развивать и совершенствовать'", "Как давно Вы используете цифровые технологии?", "Для каких целей Вы используете цифровые технологии?", "Какими программами, сайтами или приложениями Вы пользуетесь в повседневной", "Как много времени (в день) Вы уделяете на работу с цифровыми технологиями?", "Какими устройствами Вы пользуетесь?", "Укажите Ваш пол", "Напишите Ваш возраст", "Укажите Ваш уровень образования", "Как Вы оцениваете уровень своего материального положения?"]

# Создаем пустые списки для хранения данных из каждого столбца
values = [[] for _ in range(len(names))]

# Заполняем списки данными из Excel
for row in sheet.iter_rows(values_only=True):  # Итерируемся по строкам, извлекая только значения
    for i in range(len(values)):
        values[i].append(row[i+1])  # Добавляем значение из соответствующего столбца в список

# Удаляем заголовки из каждого списка
for m in range(len(values)):
    values[m] = values[m][1:]

# Создаем основное окно tkinter
root = tk.Tk()
root.title("Анализ данных")  # Устанавливаем заголовок окна

# Создаем меню
menu_bar = tk.Menu(root)
file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Выход", command=root.quit)  # Добавляем команду "Выход"
menu_bar.add_cascade(label="Файл", menu=file_menu)  # Добавляем выпадающее меню "Файл"
root.config(menu=menu_bar)  # Устанавливаем меню для окна

# Создаем метку для выбора графика
choice_label = ttk.Label(root, text="Выберите график:")
choice_label.pack(pady=5)  # Размещаем метку

# Создаем выпадающий список для выбора графика
choice_var = tk.StringVar(root)  # Создаем переменную для хранения выбранного значения
choice_dropdown = ttk.Combobox(root, textvariable=choice_var, values=[str(i+1) + ": " + name for i, name in enumerate(names)], state="readonly")  # Создаем ComboBox с названиями
choice_dropdown.pack(pady=5)  # Размещаем ComboBox
choice_dropdown.set(str(1) + ": " + names[0])  # Устанавливаем значение по умолчанию

# Создаем кнопку для отображения графика
show_button = ttk.Button(root, text="Показать график", command=lambda: show_graph(choice_var.get().split(':')[0]))  # Создаем кнопку с командой для отображения графика
show_button.pack(pady=10)  # Размещаем кнопку

# Создаем frame для размещения графика
graph_frame = tk.Frame(root)
graph_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # Размещаем frame, заполняя все доступное пространство

# Запускаем основной цикл tkinter
root.mainloop()