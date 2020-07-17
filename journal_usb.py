from tkinter import *
from tkinter import ttk, font, messagebox
from tkcalendar import DateEntry
import sqlite3
import webbrowser
from reportlab.platypus import SimpleDocTemplate
from reportlab.lib.pagesizes import A4, mm
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import inch


# вывод окошки об ошибках или информации
def msg_err(error):
    if error == 'id_usb':
        messagebox.showerror('Ошибка', 'Не указан id СНИ!')
    elif error == 'user_for_who':
        messagebox.showerror('Ошибка', 'Не указано ФИО!')
    elif error == 'user_for_who_name':
        messagebox.showerror('Ошибка', 'Указана фамилия, но без имени!')
    elif error == 'reason':
        messagebox.showerror('Ошибка', 'Не указано основание!')
    elif error == 'user_for_who_many':
        messagebox.showerror('Ошибка', 'Вероятно в ФИО указано больше значении, '
                                       '\nчем Фамилия Имя Отчество')
    elif error == 'report_completed':
        messagebox.showinfo('Информация', 'Отчет сохранен')


# создание таблицы journal_usb, если ее нет
def create_table():
    # проверяем, есть ли таблица journal_usb
    cursor.execute(''' SELECT count(name) FROM sqlite_master WHERE type='table' AND name='journal_usb' ''')
    # если нет, то создаем с указанными столбцами
    if cursor.fetchone()[0] != 1:
        cursor.execute('''CREATE TABLE journal_usb (
                        id_name INTEGER PRIMARY KEY,
                        id_usb_db TEXT, 
                        last_name_db TEXT,
                        first_name_db TEXT, 
                        middle_name TEXT,  
                        reason_db TEXT, 
                        date_db TEXT, 
                        comments_db TEXT);''')
    # сохраняем изменения в БД
    conn.commit()


# получение всех значений из таблицы
def get_from_table():
    # получаем все значения из таблицы journal_usb
    return cursor.execute('SELECT * FROM journal_usb;')


# вывод всех значении
def get_events():
    # удаляем все ранее выведенные строки
    journal_output.delete(*journal_output.get_children())
    # выводим все полученные значения из таблицы journal_usb в каждую строку
    for event in list(get_from_table()):
        journal_output.insert('', 'end', values=event)


# получение последней строки из таблицы
def get_last_added_event():
    # получаем последнюю строку из таблицы journal_usb
    last_added_event = cursor.execute('SELECT * FROM journal_usb ORDER BY id_name DESC LIMIT 1')
    # и выводим ее
    for k in last_added_event:
        journal_output.insert('', 'end', values=k)


# добавление новых значении
def add_data_to_table(*args):
    # получаем значения из вводимых значении пользователя и назначаем в переменные
    val_id_usb = id_usb.get()
    str_user_for_who = user_for_who.get()
    lst_user_for_who = str_user_for_who.split()
    val_reason = reason.get()
    val_period = period.get()
    val_comments = comments.get()

    # проверка на ошибки ввода при добавлении
    if len(val_id_usb) == 0:
        msg_err('id_usb')
        return False
    elif len(lst_user_for_who) == 0:
        msg_err('user_for_who')
        return False
    elif len(lst_user_for_who) == 1:
        msg_err('user_for_who_name')
        return False
    elif len(lst_user_for_who) > 3:
        msg_err('user_for_who_many')
        return False
    elif len(val_reason) == 0:
        msg_err('reason')
        return False

    # если в ФИО отсутствует отчество, то добавляем в отчество пустоту
    if len(lst_user_for_who) == 2:
        lst_user_for_who.append('')
    # добавляем полученные значения в таблицу
    insert_query = 'INSERT INTO journal_usb VALUES (NULL, ?, ?, ?, ?, ?, ?, ?);'
    cursor.execute(insert_query, (val_id_usb, lst_user_for_who[0], lst_user_for_who[1], lst_user_for_who[2],
                                  val_reason, val_period, val_comments))

    # сохраняем изменения в БД
    conn.commit()
    # запускаем функцию
    get_last_added_event()


def searching_in_table(*args):
    # получаем значение, что указано в поле поиск и сохраняем в переменную
    searching_val = [searching.get()]
    # удаляем весь показ журнала, чтобы в дальнейшем показать необходимое
    journal_output.delete(*journal_output.get_children())
    # проверяем какой радиобаттон выбран
    if searching_radio_sel.get() == 1:
        # получаем все данные по искомой фамилии
        select_query = 'SELECT * FROM journal_usb WHERE last_name_db=(?) ORDER BY id_name;'
    elif searching_radio_sel.get() == 2:
        # получаем все данные по искомой ID СНИ
        select_query = 'SELECT * FROM journal_usb WHERE id_usb_db=(?) ORDER BY id_name;'
    # сохраняем запрос в переменную
    execute_query = cursor.execute(select_query, searching_val)
    # выводим результат поиска
    for k in execute_query:
        journal_output.insert('', 'end', values=k)


def ask_question():
    # спрашивать пользователя, что действительно ли он хочет добавить указанные данные в БД
    window = messagebox.askquestion('Добавление данных в БД', f'Вы точно желаете добавить следующие данные в БД?\n'
                                                              f'ID СНИ\t\t{id_usb.get()}\n'
                                                              f'ФИО\t\t{user_for_who.get()}\n'
                                                              f'Основание\t{reason.get()}\n'
                                                              f'Дата\t\t{period.get()}\n'
                                                              f'Примечание\t{comments.get()}')
    # если пользователь нажмет на Да, то запускаем функцию добавления данных в БД
    if window == 'yes':
        add_data_to_table()


# создание отчета - начало ###
def chop_line(comments, max_line=13):
    # если символы в строках из столбца "Комментарии" превышает max_line, то разделяем их
    if len(comments) > max_line:
        cant = len(comments) // max_line
        cant += 1
        str_line = ''
        index = max_line
        for k in range(1, cant):
            index = max_line * k
            str_line += '%s\n' % (comments[(index - max_line):index])
        str_line += '%s' % (comments[index:])
        return str_line
    else:
        return comments


# вывод дополнительную информацию на первую страницу, а именно номер страницы
def report_first_page(canvas, doc):
    r_from = report_from.get()
    r_by = report_by.get()
    canvas.saveState()
    canvas.setFont('DejaVuSerif',16)
    canvas.drawCentredString(300, 790, 'Отчет разрешенных съемных носителей информации')
    canvas.drawCentredString(300, 770, f'с {r_from} по {r_by}')
    canvas.setFont('DejaVuSerif', 9)
    canvas.drawString(7.25 * inch, 0.75 * inch, '%d' % (doc.page))
    canvas.restoreState()


# вывод дополнительную информацию на последующие страницы, а именно номера строниц
def report_later_pages(canvas, doc):
    canvas.saveState()
    canvas.setFont('DejaVuSerif',9)
    canvas.drawString(7.25 * inch, 0.75 * inch, '%d' % (doc.page))
    canvas.restoreState()


# генерация отчета и экспорт в pdf
def create_report():
    # получаем даты отчета
    r_from = report_from.get()
    r_by = report_by.get()
    # название pdf, показывается в шапке ридера
    title = f'Отчет по разрешенным съемных носителей информации ' \
            f'\nс {r_from} по {r_by}'
    # получаем данные между казанными датами
    select_query = '''SELECT * FROM journal_usb WHERE date_db BETWEEN (?) and (?) ORDER BY id_name;'''
    # полученные данные переводим в тип списка
    list_from_table = map(list, cursor.execute(select_query, (str(r_from), str(r_by))))
    # шапка в отчете
    headers = ['ID', 'ID СНИ', 'Фамилия', 'Имя', 'Отчество', 'Основание', 'Дата', 'Примечание']
    temp_data = []
    for k in list_from_table:
        temp_data.append(k)
    # объединяем все полученные данные с шапкой
    temp_data.insert(0, headers)
    n = 0
    # тут запускаем ранее функцию для переноса не вмещяемых строк
    for k in temp_data:
        temp_data[n][7] = chop_line(k[7])
        n += 1
    # исползуем шрифт, т.к. reportlab не воспринимает кириллицу
    pdfmetrics.registerFont(TTFont('DejaVuSerif', 'DejaVuSerif.ttf', 'UTF-8'))
    # имя файла
    fileName = f'Отчет_СНИ_{r_from}_{r_by}.pdf'

    pdf = SimpleDocTemplate(
        fileName,
        pagesize=A4
    )
    pdf.title = title
    # указываем фиксированную ширину для "Комментарии"
    table = Table(temp_data, colWidths=(None, None, None, None, None, None, None, 30*mm))
    # указываем какой шрифт использовать для таблицы, т.к. если не указать, то кириллица не будет распознаваться
    style_sup_cyrillic = TableStyle([('FONTNAME', (0, 0), (-1, -1), 'DejaVuSerif')])
    table.setStyle(style_sup_cyrillic)
    # дальше стилистика таблицы
    style_header = TableStyle([
        ('BACKGROUND', (0, 0), (8, 0), colors.fidblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
    ])
    table.setStyle(style_header)

    rowNumb = len(temp_data)
    for i in range(1, rowNumb):
        if i % 2 == 0:
            bc = colors.burlywood
        else:
            bc = colors.beige
        ts = TableStyle(
            [('BACKGROUND', (0, i), (-1, i), bc)]
        )
        table.setStyle(ts)

    ts = TableStyle(
        [
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]
    )
    table.setStyle(ts)
    # создаем список значении
    elems = []
    elems.append(table)
    # создаем pdf, где добавляем таблицу и шапку для первой страницы и нумерацию для первой и последующие страницы
    pdf.build(elems, onFirstPage=report_first_page, onLaterPages=report_later_pages)
    msg_err('report_completed')
# создание отчета - конец ###


# интерфейс - начало ###
root = Tk()
root.title('Журнал разрешенных СНИ')
root.resizable(width=False, height=False)

first_frame = ttk.Frame(root, borderwidth=5)
first_frame.grid(column=0, row=0, sticky=(N, W, E, S))

font_bold = font.Font(weight='bold', size=10)

ttk.Label(first_frame, text='ID СНИ', font=font_bold) \
    .grid(column=1, row=0, sticky=W)
ttk.Label(first_frame, text='ФИО - Кому', font=font_bold) \
    .grid(column=3, row=0, sticky=W)
ttk.Label(first_frame, text='Основание', font=font_bold) \
    .grid(column=5, row=0, sticky=W)
ttk.Label(first_frame, text='Дата', font=font_bold) \
    .grid(column=7, row=0, sticky=W)
ttk.Label(first_frame, text='Примечание', font=font_bold) \
    .grid(column=9, row=0, sticky=W)

id_usb = ttk.Entry(first_frame, width=25)
id_usb.grid(column=1, row=1, sticky=W)
user_for_who = ttk.Entry(first_frame, width=35)
user_for_who.grid(column=3, row=1, sticky=W)
reason = ttk.Entry(first_frame, width=15)
reason.grid(column=5, row=1, sticky=W)
period = DateEntry(first_frame, date_pattern='YYYY-MM-DD', width=10)
period.grid(column=7, row=1, sticky=W)
comments = ttk.Entry(first_frame, width=35)
comments.grid(column=9, row=1, sticky=W)
ttk.Button(first_frame, text='Добавить', command=ask_question) \
    .grid(column=11, row=1, sticky=W)

# добавление пустое место между столбцами
first_frame.grid_columnconfigure(2, minsize=10)
first_frame.grid_columnconfigure(4, minsize=10)
first_frame.grid_columnconfigure(6, minsize=10)
first_frame.grid_columnconfigure(8, minsize=10)
first_frame.grid_columnconfigure(10, minsize=10)

second_frame = ttk.Frame(root, borderwidth=5)
second_frame.grid(column=0, row=1, sticky=(N, W, E, S))

ttk.Label(second_frame, text='Поиск', font=font_bold) \
    .grid(column=1, row=1, sticky=W)
searching = ttk.Entry(second_frame, width=25)
searching.grid(column=1, row=2, sticky=W)
# следующая строка необходима, чтобы потом из ниже радиобаттнов сделать необходимую по умолчанию
searching_radio_sel = IntVar()
searching_radio_name = ttk.Radiobutton(second_frame, text='Фамилия', value=1, variable=searching_radio_sel)
searching_radio_name.grid(column=2, row=2, sticky=W)
searching_radio_idusb = ttk.Radiobutton(second_frame, text='ID СНИ', value=2, variable=searching_radio_sel)
searching_radio_idusb.grid(column=3, row=2, sticky=W)
# вот тут и указываем, как из радиобаттнов будет по умолчанию
searching_radio_sel.set(1)
ttk.Button(second_frame, text='Найти', command=searching_in_table) \
    .grid(column=4, row=2, sticky=W)
ttk.Button(second_frame, text='Сброс', command=get_events) \
    .grid(column=5, row=2, sticky=W)

ttk.Label(second_frame, text='Отчет', font=font_bold) \
    .grid(column=7, row=1, sticky=W)

report_from = DateEntry(second_frame, date_pattern='YYYY-MM-DD', width=10)
report_from.grid(column=7, row=2, sticky=W)
report_by = DateEntry(second_frame, date_pattern='YYYY-MM-DD', width=10)
report_by.grid(column=9, row=2, sticky=W)

ttk.Button(second_frame, text='Экспортировать', command=create_report) \
    .grid(column=11, row=2, sticky=W)
second_frame.grid_columnconfigure(6, minsize=178)

second_frame.grid_columnconfigure(8, minsize=3)
second_frame.grid_columnconfigure(10, minsize=3)

third_frame = ttk.Frame(root, borderwidth=5)
third_frame.grid(column=0, row=3, sticky=(N, W, E, S))

ttk.Label(third_frame, text='Журнал', font=font_bold) \
    .grid(column=1, row=0, sticky=W)

cols = ('ID', 'ID СНИ', 'Фамилия', 'Имя', 'Отчество', 'Основание', 'Дата', 'Примечание')
journal_output = ttk.Treeview(third_frame, columns=cols, show='headings', height=30)
journal_output.heading('ID', text='ID')
journal_output.column('ID', width=50)
journal_output.heading('ID СНИ', text='ID СНИ')
journal_output.column('ID СНИ', width=110)
journal_output.heading('Фамилия', text='Фамилия')
journal_output.column('Фамилия', width=110)
journal_output.heading('Имя', text='Имя')
journal_output.column('Имя', width=110)
journal_output.heading('Отчество', text='Отчество')
journal_output.column('Отчество', width=110)
journal_output.heading('Основание', text='Основание')
journal_output.column('Основание', width=100)
journal_output.heading('Дата', text='Дата')
journal_output.column('Дата', width=70)
journal_output.heading('Примечание', text='Примечание')
journal_output.column('Примечание', width=220)
journal_output.grid(column=1, row=1)

scroll_y = ttk.Scrollbar(third_frame, orient="vertical", command=journal_output.yview)
scroll_y.grid(column=2, row=1, sticky=(N, S))
scroll_x = ttk.Scrollbar(third_frame, orient="horizontal", command=journal_output.xview)
scroll_x.grid(column=1, row=2, sticky=(W, E))

author = ttk.Label(third_frame, text='daradan')
author.grid(column=1, row=3, sticky=W)
author.bind('<Button-1>', lambda event: webbrowser.open("https://github.com/daradan?tab=repositories"))

# donate = ttk.Label(third_frame, text='donate')
# donate.grid(column=1, row=3, sticky=W)
# donate.bind('<Button-1>', lambda event: webbrowser.open("https://www.paypal.me/daradan?locale.x=ru_RU"))

version = ttk.Label(third_frame, text='v0.1(20200717)')
version.grid(column=1, row=3, sticky=E, columnspan=2)
# интерфейс - конец ###

conn = sqlite3.connect('oib.db')
cursor = conn.cursor()

create_table()

get_events()

id_usb.focus()

root.mainloop()
