from vk_api.bot_longpoll import VkBotLongPoll
from vk_api.bot_longpoll import VkBotEventType
from PIL import Image, ImageDraw, ImageFont
from keys import token_vk, id_group_vk
from threading import Thread

import traceback
import requests
import vk_api
import inst
import xlrd

'''Раздел автризации'''
vk = vk_api.VkApi(token=token_vk)
session_api = vk.get_api()
longPool = VkBotLongPoll(vk, id_group_vk)

# Отправка сообщения о успешном запуске в личные сообщения вк
vk.method('messages.send', {'peer_id': 166799901, 'random_id': 0, 'message': 'Бот успешно запущен!'})


class MyError(Exception):
    """
    Класс для реализации собственных ошибок (Для возврата ошибки пользователю при
    неверно введёных данных)
    """
    pass


class GetXLSX(Thread):
    """
    Класс для обновления таблиц-расписаний
    __inti__ - Конструктор клаасса
        self.chat_id - Переменная, хранящаяя ID чата, куда необходимо отправить сообщение

    run - Функция ассинхронного действия. Выполняет действия "отдельно" от кода
          Парсит сайт с расписанием, скачивая необходимые файлы, затем обновляет словарь с ключами-названиями файлов
          и значениями-группы, которые находятся в этом файле.
    """
    def __init__(self, id_chat):
        Thread.__init__(self)
        self.chat_id = id_chat  # ID чата в вк

    def run(self):
        try:
            a = requests.get('https://www.mirea.ru/schedule/').text  # Веб-запрос с получением HTML кода страницы
            b = [i.split('"')[-1] + '.xlsx' for i in a.split('.xlsx')][:-1]  # Список ссылок на файлы XLSX
            new_inst = {}  # Словарь для хранения файл-группы
            xlsx, gr = 0, 0  # Переменные для статистики обработанной информации (Кол-во таблиц и групп)
            for i in b:
                r = requests.get(i)  # Веб-запрос с получением таблицы
                with open('tables/' + i.split('/')[-1], "wb") as code:  # Запись таблицы
                    code.write(r.content)
                table = xlrd.open_workbook('tables/' + i.split('/')[-1])  # Открытие таблицы
                sheet = table.sheet_by_index(0)  # Выбор 1 листа
                vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]  # Запись данных в список списков
                # Список всех групп текущего файла
                group = [vals[1][i].split()[0] for i in range(len(vals[1])) if str(vals[1][i]).count('-') == 2]
                new_inst[i.split('/')[-1]] = group  # Запись в словарь файла и групп
                xlsx += 1  # Обновляем количество обработанных файлов
                gr += len(group)  # Обновляем количество обработанных групп
            inst.update_inst(new_inst)  # Обновление словаря файл-группы для дальнейшей работы с ней
            # Отпправка сообщения о успешном скачивании и установке
            vk.method('messages.send', {'peer_id': self.chat_id, 'random_id': 0,
                                        'message': '✅Успешных установок XLMS - {}, Групп найдено - {}'.format(
                                            xlsx, gr)})
        except Exception:  # При возникновении любой ошибки сообщение о ней отправится в личные сообщения разработчика
            vk.method('messages.send', {'peer_id': self.chat_id, 'random_id': 0,
                                        'message': 'Опс... Похоже произошёл сбой, попробуйте позже...'})
            vk.method('messages.send', {'peer_id': 166799901, 'message': traceback.format_exc(), 'random_id': 0})


class GetTimetable(Thread):
    """
    Класс для создания картинки с расписанием
    __init__ - конструктор класса
        self.chat_id - Хранит ID чата в ВК
        self.group - Хранит название необходимой группы
        self.chet_image - Хранит номер картинки (Для удобного сохранения в последующем)
        self.shablon_id - Хранит номер выбраного дизайна шаблона
        self.design - Хранит словарь с дизайнами
    run - Функция ассинхронного действия. Выполняет действия "отдельно" от кода
          Отрисовывает изображение согласно выбраному дизайну и группе
    """
    def __init__(self, id_chat, name_group, image_chet, id_shablon=4):
        Thread.__init__(self)
        self.chat_id = id_chat
        self.group = name_group.upper()
        self.chet_image = image_chet
        self.shablon_id = id_shablon
        self.design = {1: ['shablons/shablon1.png', (255, 255, 255, 255)],
                       2: ['shablons/shablon2.png', (255, 255, 255, 255)],
                       3: ['shablons/shablon3.png', (0, 0, 0, 255)],
                       4: ['shablons/shablon4.png', (0, 0, 0, 255)]}

    def run(self):
        try:
            name_xlsx = ''  # Переменная для хранения названия необходимой таблицы с расписанием
            for k, v in inst.inst.items():  # Проход по ключам и значениям словаря файл-группы
                if self.group in v:  # Если нужная группа находится в группах текущего файла
                    name_xlsx = k  # Обновляем название файла на ключ словаря
                    break  # Завершаем цикл дострочно, поскольку файл найден
            if name_xlsx == '':  # Если файл не найден
                raise MyError("Неверно выбрана группа!")  # Вернёт ошибку
            institut, kurs = '', ''  # Переменные для хранения названия института и курса
            # Название института и курс указаны в название файла, поэтому выделим их от туда
            for i in name_xlsx:
                if i.upper() == i:
                    institut += i
                elif i != '':
                    break
            for i in name_xlsx:
                if i.isdigit():
                    kurs = i
                    break
            rb = xlrd.open_workbook('tables/{}'.format(name_xlsx))  # Открывает необходимую таблицу
            sheet = rb.sheet_by_index(0)  # Выбираем 1 лист
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]  # Создаем список списков из таблицы
            # Находим номер столбца с необходимой нам группой
            group = [i for i in range(len(vals[1])) if vals[1][i] == self.group][0]
            # Создаем список списков с расписанием нужной группы [Предмет, кабинет, вид преподования]
            pars = [[vals[i][group], vals[i][group + 1], vals[i][group + 3]] for i in range(3, 75)]
            # Создание изменяемого изображения на основе выбраного шаблона
            im = Image.open(self.design[self.shablon_id][0])
            # Создание объекта для отрисовки данных
            draw = ImageDraw.Draw(im)
            # Инициализация необходимых шрифтов для текстов
            subject_fnt = ImageFont.truetype('fonts/msyh.ttc', 14)
            title_fnt = ImageFont.truetype('fonts/msyh.ttc', 24)
            # Отрисовка группы, института и курса соответственно
            draw.text((119, 23), self.group, font=title_fnt, fill=self.design[self.shablon_id][1])
            draw.text((1082, 23), institut, font=title_fnt, fill=self.design[self.shablon_id][1])
            draw.text((1272, 23), kurs, font=title_fnt, fill=self.design[self.shablon_id][1])
            # Координаты начала написания текста расписания
            x = 132
            y = 128
            for i in range(len(pars)):
                # Отрисовка названия предмета, номера кабинета и вида преподования соотвественно
                draw.text((x, y), pars[i][0][:20].strip() + '.' if len(pars[i][0]) > 20 else pars[i][0],
                          font=subject_fnt, fill=self.design[self.shablon_id][1])
                draw.text((x, y + 20), pars[i][2].replace('\n', '/'),
                          font=subject_fnt, fill=self.design[self.shablon_id][1])
                w, h = draw.textsize(pars[i][1].upper().replace('\n', '/'), font=subject_fnt)
                draw.text((x + 168 - w, y + 20), pars[i][1].upper().replace('\n', '/'),
                          font=subject_fnt, fill=self.design[self.shablon_id][1])
                # При достижении определённой высоты нам необходимо перейти в следующий столбец
                if y == 689:
                    y = 128
                    x += 191
                else:
                    y += 51
            im.save('ras{}.png'.format(self.chet_image))  # Сохраняем изображение
            # Загрузка изображения в вк для последующей его отправки
            met = vk.method('photos.getMessagesUploadServer')
            b = requests.post(met['upload_url'],
                              files={'photo': open('ras{}.png'.format(self.chet_image), 'rb')}).json()
            c = vk.method('photos.saveMessagesPhoto',
                          {'photo': b['photo'], 'server': b['server'], 'hash': b['hash']})[0]
            d = 'photo{}_{}'.format(c['owner_id'], c['id'])
            # Отправка сообщения с картинкой
            vk.method('messages.send', {'peer_id': self.chat_id, 'attachment': d, 'random_id': 0})
        except MyError as e:  # Отправка сообщения пользователю при ошибке с его стороны
            vk.method('messages.send', {'peer_id': self.chat_id, 'random_id': 0, 'message': "Опс... %s" % e})
        except Exception:  # Отправка сообщения разработчику при ощибке со стороны кода
            vk.method('messages.send', {'peer_id': self.chat_id, 'random_id': 0,
                                        'message': 'Опс... Похоже произошёл сбой, попробуйте позже...'})
            vk.method('messages.send', {'peer_id': 166799901, 'message': traceback.format_exc(), 'random_id': 0})


chet_image = 0  # Отвечает за счет сохраняемых изображений
running = True   # Костыль, для того, чтобы компилятор не выделял бесконечный цикл
while running:
    try:
        for event in longPool.listen():  # Проходимся по всем ивентам, которые бот отследил
            if event.type == VkBotEventType.MESSAGE_NEW:  # Если ивент - это пришедшее собщение
                chat_id = event.object.peer_id  # ID Чата, откуда пришло сообщение
                from_id = event.object.from_id  # ID Пользователя, от которого пришло сообщение
                text_m = event.object.text.lower()  # Текст сообщения
                if len(text_m) > 4:  # Постое отсеивание ненужных сообщений по длине
                    if len(text_m.split('-')) == 3:  # Если в сообщении всего 2 знака '-'
                        if len(text_m.split('-')[1]) == 1:  # Если номер группы указан в виде одного числа, а не двух
                            a = text_m.split('-')
                            text_m = a[0] + '-0' + a[1] + '-' + a[2]  # Редактируем номер группы под необходимый формат
                        if len(text_m.split()) == 2 and text_m.split()[1] in ['1', '2', '3', '4']:  # Если выбран дизайн
                            timetable = GetTimetable(chat_id, text_m.split()[0], chet_image, int(text_m.split()[1]))
                            timetable.start()
                            chet_image += 1
                            chet_image %= 10
                        else:
                            timetable = GetTimetable(chat_id, text_m.split()[0], chet_image)
                            timetable.start()
                            chet_image += 1
                            chet_image %= 10
                    elif text_m == 'обновить' and from_id == 166799901:  # Обновление данных, доступ только у создателя
                        obn = GetXLSX(chat_id)
                        obn.start()
    except Exception:
        pass
