import PySimpleGUI as sg # pip install PySimpleGUI
import sqlite3 # pip install pysqlite3
import re # Для проверки правильности емайла
from fpdf import FPDF # Создание PDF
# Библиотеки используемые при отправке почты
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Иницализиурем соединения с базой данных
database = sqlite3.connect("schedule.db")
cursor = database.cursor()

FREE_TIME_NAME = "свободный час"
DAYS = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье']
PDF_FILE = "schedule.pdf"

# Создаем таблицы если база пустая
cursor.execute('''CREATE TABLE IF NOT EXISTS logins (Name text PRIMARY KEY, Surname text, Login text, Password text, Role text)''')
database.commit()
cursor.execute('''CREATE TABLE IF NOT EXISTS subjs (Id integer PRIMARY KEY AUTOINCREMENT, Subj text)''')
database.commit()
# Добавляем пропуск в расписание для удобства
data = cursor.execute('''SELECT * FROM subjs''')
if not data.fetchall():
    data = cursor.execute('''INSERT INTO subjs (Subj) values (?)''', (FREE_TIME_NAME,))
    database.commit()
# В таблице schedule будет Внешний Ключ связывающий ее с таблицей предметов subjs
cursor.execute('''CREATE TABLE IF NOT EXISTS schedule (Day text, Time integer, SubjId integer, FOREIGN KEY (SubjId) REFERENCES subjs (Id))''')
database.commit()


# Функция добавления нового пользователя
def insert_user(name, surname, login, password, role):
    cursor.execute('''INSERT INTO logins (Name,Surname,Login,Password,Role) values (?,?,?,?,?)''',
                        (name, surname, login, password, role))
    database.commit()


# Функция поиска пользователя в БД
def get_user(login):
    data = cursor.execute('''SELECT * FROM logins WHERE Login = ?''', (login,))
    return data.fetchone()


# Наименование предмета по его id
def get_subj_byname(name):
    data = cursor.execute('''SELECT Id FROM subjs WHERE Subj = ?''', (name,))
    return data.fetchone()[0]

# Функция получения списка предметов из БД
def get_subjs():
    data = cursor.execute('''SELECT * FROM subjs''')
    return data.fetchall()


# Функция добавления предмета в БД
def add_subj(name):
    cursor.execute('''INSERT INTO subjs (Subj) values (?)''', (name,))
    database.commit()


# Функция редактирования предмета в БД
def update_subj(name, id):
    cursor.execute('''UPDATE subjs SET Subj = ? WHERE Id = ?''', (name, id))
    database.commit()


# Функция удаления предмета в БД
def delete_subj(id):
    cursor.execute('''DELETE FROM subjs WHERE Id = ?''', (id, ))
    database.commit()


# Расписание дня (отсортированное)
def get_schedule(day):
    data = cursor.execute('''SELECT schedule.Time, subjs.Subj FROM schedule, subjs WHERE Day = ? AND schedule.SubjId = subjs.Id ORDER BY schedule.Time ASC''', (day,))
    return data.fetchall()


# Расписание дня (отсортированное в обратном порядке)
def get_schedule_desc(day):
    data = cursor.execute('''SELECT Time, SubjId FROM schedule WHERE Day = ? ORDER BY Time DESC''', (day,))
    return data.fetchall()


# Предмет по дню и времени
def get_schedule_time(day, time):
    data = cursor.execute('''SELECT subjs.Subj FROM schedule, subjs WHERE Day = ? AND Time = ? AND schedule.SubjId = subjs.Id''', (day, time))
    return data.fetchall()


# Добавить предмет в расписание дня
def add_schedule(day, time, subj_id):
    cursor.execute('''INSERT INTO schedule (Day, Time, SubjId) values (?,?,?)''', (day, time, subj_id))
    database.commit()


# Исправить предмет в расписании
def update_schedule(day, time, subj_id):
    if get_schedule_time(day, time):
        cursor.execute('''UPDATE schedule SET SubjId = ? WHERE Day = ? AND Time = ?''', (subj_id, day, time))
        database.commit()
    else:
        add_schedule(day, time, subj_id)


# Функция удаления предмета из расписания
def delete_schedule(day, time):
    cursor.execute('''DELETE FROM schedule WHERE Day = ? AND Time = ?''', (day, time))
    database.commit()


# Окно логина
def login_window():
    layout = [
        [sg.Text('Введите Логин/Пароль')],
        [sg.Text('')],
        [sg.Text('Логин:   ', font=("Helvetica", 12)), sg.InputText('', size=(15, 1))],
        [sg.Text('Пароль:', font=("Helvetica", 12)), sg.InputText('', size=(15, 1))],
        [sg.Text('')],
        [sg.Submit('Вход'), sg.Text('        '), sg.Submit('Регистрация')]
    ]

    window = sg.Window('Авторизация', layout)
    while True:
        event, values = window.read()

        if event in (None, sg.WIN_CLOSED):
            window.close()
            return None

        if event == 'Вход':
            login = values[0]
            pasw = values[1]
            if len(login) == 0 or len(pasw) == 0:
                sg.popup_error('Поля логина и пароля не должны быть пустыми')
                continue
            user = get_user(values[0])
            if not user:
                sg.popup_error('Пользователя не существует')
                continue
            if pasw == user[3]:
                window.close()
                return(user)
            else:
                sg.popup_error('Неверный пароль')

        if event == 'Регистрация':
            register_window()


# Окно регистрации пользователя
def register_window():
    layout = [
        [sg.Text('Поля * обязательны для заполнения')],
        [sg.Text('')],
        [sg.Text('Логин*:    ', font=("Helvetica", 12)), sg.InputText('', size=(22, 1))],
        [sg.Text('Пароль*: ', font=("Helvetica", 12)), sg.InputText('', size=(22, 1))],
        [sg.Text('Имя*:       ', font=("Helvetica", 12)), sg.InputText('', size=(22, 1))],
        [sg.Text('Фамилия:', font=("Helvetica", 12)), sg.InputText('', size=(22, 1))],
        [sg.Text('Права*:   ', font=("Helvetica", 12)), sg.InputCombo(('Гость', 'Админ'), size=(20, 1))],
        [sg.Text('')],
        [sg.Submit('Регистрация')]
    ]

    window = sg.Window('Регистрация', layout)
    while True:
        event, values = window.read()

        if event in (None, sg.WIN_CLOSED):
            window.close()
            return None

        if event == 'Регистрация':
            login = values[0]
            pasw = values[1]
            name = values[2]
            sname = values[3]
            role = values[4]
            if len(login) == 0 or len(pasw) == 0 or len(name) == 0 or len(role) == 0:
                sg.popup_error('Поля * не должны быть пустыми')
                continue
            user = get_user(login)
            if user:
                sg.popup_error('Пользователь существует')
                continue
            else:
                insert_user(name, sname, login, pasw, role)
                window.close()


# Дополнительная функция описания прав пользователя
def role_text(role):
    text = "Вы можете: "
    if role == "Гость":
        text += "Просматривать расписание без редактирования"
    if role == "Админ":
        text += "Просматривать и редактировать расписание"
    return text


# Доп фукция удаления пустых хвостов в расписании
def clear_tail(day):
    schedule = get_schedule_desc(day)
    for elem in schedule:
        if elem[1] == 1:
            delete_schedule(day, elem[0])
        else:
            break


# Главное окно работы с расписанием
def main_window(user):
    subj_list = []
    subjs = get_subjs()
    for subj in subjs:
        subj_list.append(subj[1])
    layout = [
        [sg.Text('Добро пожаловать, %s!' % user[0], size=(25, 1), justification='center', font=("Helvetica", 20))],
        [sg.Text('Ваш уровень прав в программе: %s' % user[4])],
        [sg.Text(role_text(user[4]))]
    ]

    for day in DAYS:
        layout.append([sg.Text('%s' % day, size=(25, 1), justification='center', font=("Helvetica", 14))])
        schedule = get_schedule(day)
        i = 0
        for subj in schedule:
            num = subj[0] + 1
            layout.append([sg.Text('%i' % num, size=(10, 1)), sg.InputCombo(subj_list, default_value=subj[1], key=(day, subj[0]), size=(30, 1))])
            i += 1
        layout.append([sg.Text('Добавить - ', size=(10, 1)), sg.InputCombo(subj_list, key=(day, i), size=(30, 1))])

    layout.extend([[sg.Text('')],
        [sg.Submit('Сохранить'), sg.Text(''), sg.Submit('Сохранить как PDF'), sg.Text(''), sg.Submit('Редактировать предметы')],
        [sg.Submit('Отправить на почту - '), sg.InputText('', size=(36, 1), key='email')]])

    window = sg.Window('Просмотр и редактирование расписания', layout)
    while True:
        event, values = window.read()
        if event in (None, sg.WIN_CLOSED):
            window.close()
            exit(0)
        if event == 'Сохранить':
            if user[4] == "Гость":
                sg.popup_error('Не достаточно прав')
            else:
                window.close()
                save_subjs(values)
                main_window(user)
                sg.popup('Расписание сохранено')
        if event == 'Редактировать предметы':
            save_subjs(values)
            window.close()
            subj_window(user)

        if event == 'Сохранить как PDF':
            save_pdf()
            sg.popup('Расписание сохранено в файл %s' % PDF_FILE)

        if event == 'Отправить на почту - ':
            values['email'] = values['email'].strip()
            if len(values['email']) == 0 or not re.match(r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)", values['email']):
                sg.popup_error('Введите корректный адрес почту')
            else:
                send_email(values)


# Окно редактирования предметов
def subj_window(user):
    layout = [
        [sg.Text('Изменить существующие названия')],
        [sg.Text('Или добавить новый предмет')]
    ]
    subjs = get_subjs()
    for subj in subjs:
        layout.append([sg.Text(subj[0], size=(5, 1)), sg.InputText(subj[1], size=(23, 1), key=subj[0])])
    layout.append([sg.Text('Новый', size=(5, 1)), sg.InputText('', size=(23, 1), key='new')])
    layout.append([sg.Submit('Изменить'), sg.Text('        '), sg.Submit('К расписанию')])

    window2 = sg.Window('Просмотр и редактирование расписания', layout)
    while True:
        event, values = window2.read()
        if event in (None, sg.WIN_CLOSED):
            window2.close()
            main_window(user)
        if event == 'Изменить':
            for key in values:
                if key == 'new':
                    if len(values['new']):
                        add_subj(values['new'])
                else:
                    if len(values[key]) == 0:
                        delete_subj(key)
                    else:
                        update_subj(values[key], key)
            window2.close()
            subj_window(user)
        if event == 'К расписанию':
            window2.close()
            main_window(user)


def save_subjs(values):
    #print(values)
    for key in values:
        if not len (values[key]):
            subj_id = 1
        else:
            subj_id = get_subj_byname(values[key])
        if key != 'email':
            update_schedule(key[0], key[1], subj_id)
    for day in DAYS:
        clear_tail(day)


def send_email(values):
    HOST = "smtp.mail.ru"
    SUBJECT = "Расписание занятий"
    FROM = "vebinariumlanguages@mail.ru"
    FROM_PSWD = "MarksTeamBestik"
    TEXT = "РАСПИСАНИЕ\n\n"

    for day in DAYS:
        TEXT += day + '\n'
        schedule = get_schedule(day)
        if not schedule:
            TEXT += "Нет занятий\n"
        for subj in schedule:
            num = subj[0] + 1
            TEXT += "%i - %s\n" % (num, subj[1])
        TEXT += '\n'

    msg = MIMEMultipart()
    msg['From'] = FROM
    msg['To'] = values['email']
    msg['Subject'] = SUBJECT
    body = TEXT
    msg.attach(MIMEText(body, 'plain'))
    server = smtplib.SMTP_SSL(HOST, 465)
    server.login(FROM, FROM_PSWD)
    server.send_message(msg)
    server.quit()
    sg.popup('Расписание отправлено по адресу %s' % values['email'])


def save_pdf():
    pdf = FPDF(orientation='P', unit='mm', format='A4')

    pdf.add_font('GNU FreeFont', '', '9807.ttf', uni=True)
    pdf.add_page()
    pdf.set_font('GNU FreeFont', '', 16)
    pdf.cell(190, 15, txt="РАСПИСАНИЕ", ln=1, align="C")
    for day in DAYS:
        pdf.set_font('GNU FreeFont', '', 14)
        pdf.cell(190, 15, txt=day, ln=1, align="C")
        pdf.set_font('GNU FreeFont', '', 12)
        schedule = get_schedule(day)
        if not schedule:
            pdf.cell(190, 10, txt="Нет занятий", ln=1, align="L")
        for subj in schedule:
            num = subj[0] + 1
            pdf.cell(190, 10, txt="%i - %s" % (num, subj[1]), ln=1, align="L")

    pdf.output(PDF_FILE)


def main():
    user = login_window()
    if user:
        main_window(user)


if __name__ == "__main__":
    main()
