import telebot
from selenium import webdriver
from selenium.webdriver.common.by import By
import xlwt

bot = telebot.TeleBot("2087589116:AAGY4XjzWLkfz6MiRdxLkGhBpY_qOFIcB7g")


keyboard = telebot.types.ReplyKeyboardMarkup(True)
keyboard.row('/start')


@bot.message_handler(commands=['start'])
def start_message(message):
    bot.send_message(message.chat.id, 'Привіт! Щоб отримати звіт про заміни введіть наступні дані:')
    sent = bot.send_message(message.chat.id, 'ПІБ повністю')
    bot.register_next_step_handler(sent, names)


def names(message):
    name = message.text
    # Открываем файл для записи
    file = open('n' + message.from_user.first_name + '.txt', 'w')
    # Записываем
    file.write(name)
    # Закрываем файл
    file.close()
    sent1 = bot.send_message(message.chat.id, 'Початкова дата у форматі дд.мм.рррр (01.09.2021)')
    bot.register_next_step_handler(sent1, pochs)


def pochs(message):
    poch = message.text
    # Открываем файл для записи
    file = open('p' + message.from_user.first_name + '.txt', 'w')
    # Записываем
    file.write(poch)
    # Закрываем файл
    file.close()
    sent2 = bot.send_message(message.chat.id, 'Кінцева дата у форматі дд.мм.рррр (30.09.2021)')
    bot.register_next_step_handler(sent2, kincs)


def kincs(message):
    kinc = message.text
    bot.send_message(message.chat.id, 'Йде обробка зачекайте...')

    with open('n' + message.from_user.first_name + '.txt') as file:
        name = file.read()
    with open('p' + message.from_user.first_name + '.txt') as file1:
        poch = file1.read()

    # Відкрити хром
    driver = webdriver.Chrome()
    driver.get('https://dekanat.nung.edu.ua/cgi-bin/timetable.cgi')

    # Ввести значення у відповідності до імені
    element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/form/div[2]/div[1]/input')
    element.send_keys(name)

    # Ввести початкову дату
    element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/form/div[3]/div[1]/div/div/input')
    element.send_keys(poch)

    # Ввести кінцеву дату
    element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/form/div[3]/div[2]/div/div/input')
    element.send_keys(kinc)

    # Натиснути на кнопку "Показати розклад"
    button = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/form/div[3]/div[3]/button')
    button.click()

    pos = driver.find_elements(By.XPATH, 'col-md-6')

    # Ініціалізація робочої книги
    book = xlwt.Workbook(encoding="utf-8")

    # Додавання аркушу в книгу
    sheet1 = book.add_sheet("Дні")

    # Створення заголовків
    sheet1.write(0, 0, '№ запису')
    sheet1.write(0, 1, 'Дата')
    sheet1.write(0, 2, 'Пара')
    sheet1.write(0, 3, 'Початок')
    sheet1.write(0, 4, 'Кінець')
    sheet1.write(0, 5, 'Тип')
    sheet1.write(0, 6, 'Назва заняття')
    sheet1.write(0, 7, 'Групи')

    # Пошук і запис замін
    # Ім'я файлу
    namef = 'Заміни_' + message.from_user.first_name + '_' + poch + '_' + kinc + '.xls'
    filename = namef

    find = "заміні"
    i = 1
    elements = driver.find_elements(By.CLASS_NAME, 'col-md-6')
    for class1 in elements:
        textcl = class1.text
        if find in textcl:
            text = textcl.split()
            sheet1.write(i, 1, text[0])
            i = i + textcl.count(find)

    i = 1
    elements1 = driver.find_elements(By.TAG_NAME, 'tr')
    for tr in elements1:
        texttr = tr.text
        if find in texttr:
            text1 = texttr.split()
            sheet1.write(i, 2, text1[0])
            sheet1.write(i, 3, text1[1])
            sheet1.write(i, 4, text1[2])
            i = i + 1

    i = 1
    elements2 = driver.find_elements(By.TAG_NAME, 'td')
    for td in elements2:
        texttd = td.text
        if find in texttd:
            text2 = texttd.split("\n")
            sheet1.write(i, 0, i)
            sheet1.write(i, 5, text2[0])
            sheet1.write(i, 6, text2[1])
            sheet1.write(i, 7, text2[2])
            i = i + 1


    # Збереження файлу
    book.save(filename)

    driver.quit()

    with open(namef, "rb") as f:
        bot.send_document(message.chat.id, f, caption="Ваш файл")


bot.infinity_polling()
