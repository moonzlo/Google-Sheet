#!/venv/bin/ python

from oauth2client.service_account import ServiceAccountCredentials
from func import *

from multiprocessing.dummy import Pool
import gc



start = time.time()
def get_html(url):
    agent = 'Mozilla/5.0 (Windows NT 6.0) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.52 Safari/536.5'
    headers = {'accept': '*/*',
               'user-agent': agent}
    session = requests.Session()  # Иметирует сессию клиента.
    request = session.get(url, headers=headers)


    return request.content


def get_sheet():
    '''отдаёт список с словарями, количество словарей равно количесту заполненных строк'''

    scope = ['https://www.googleapis.com/auth/drive']

    creds = ServiceAccountCredentials.from_json_keyfile_name('login.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open('test1').sheet1

    leg = sheet.get_all_records()  # Список внутри которого словари (количество словареий равно кличеству строк)
    mass = leg[4:]  # Пропускаем пустые строки

    return sheet


def reed_json():
    with open('test.json', 'r', encoding='utf-8') as js:  # открываем файл на чтение
        data = json.load(js)  # загружаем из файла данные в словарь data

    return data


class Stroka(object):

    def __init__(self, dictionary, deck):
        '''Принимает словарь из списка открытого json файла, задает базовые значения классу.'''


        self.sheet = deck  # Объект доски.
        self.time = dictionary.get('Отметка времени')
        self.name = dictionary.get('Фамилия и имя')
        self.doing = dictionary.get('Дейсвтие')
        if self.doing == None:
            self.doing = ''
        self.items_name = dictionary.get('Наименование (кратко)')
        self.article = dictionary.get('Артикул')
        self.item_url = dictionary.get('Ссылка на товар')
        self.unit_item = dictionary.get('Цена за единицу товара (в рос.руб)')
        if self.unit_item == None:
            self.unit_item = ''
        self.item_value = dictionary.get('Количество (нужное Вам)')
        self.min_value = dictionary.get('Количество для минимального заказа')
        self.notice = dictionary.get('Примечание')
        self.delivery = dictionary.get('Цена за доставку (если есть)')
        self.str_number = dictionary.get('Строка')

        # self.klaster = dictionary.get(keys[13])  # Параметр для заглушки, хранит имя кластера


    def table_update(self):
        '''Данный метод записывает данные непосредственно в таблицу'''

        namers = [self.time, self.name, self.doing, self.items_name, self.article, self.item_url, self.unit_item,
                  self.item_value, self.min_value, self.delivery, '', '', self.notice]

        if bool(self.name) != False:

            cell_list = self.sheet.range('A{}:M{}'.format(self.str_number, self.str_number))

            # Обновляем каждоу строку.
            for cell, name in zip(cell_list, namers):
                cell.value = name

            # Update in batch
            self.sheet.update_cells(cell_list)



    def product_name_update(self):
        '''Метод получает полное название + цену товара, и обновляет их в параметрах экземпляра класса'''

        base = 'https://www.sima-land.ru/{}'.format(self.article)
        html = get_html(base)

        data = sopchik(html)
        table_numbers = 'A{}:M{}'.format(self.str_number, self.str_number)

        if data[0] == True:
            # Если тэг Товар партнёра Найден, окрашивает в берюзовый всю строку.
            default_format = CellFormat(backgroundColor=color(250, 10, 10), textFormat=textFormat(bold=False))
            format_cell_range(self.sheet, table_numbers, default_format)


        # Обновляем название и цену
        try:
            price = data[1]

            a = re.findall(r'\d+', '{}'.format(price))
            num = str().join(a)
            self.unit_item = int(num)

            price_name = data[2]
            for i in price_name:
                self.items_name = i



            status = data[3]
            if status == 'Нет в наличии':
                default_format = CellFormat(backgroundColor=color(1, 0, 240), textFormat=textFormat(bold=False))
                format_cell_range(self.sheet, table_numbers, default_format)


        except AttributeError:
            self.unit_item = 'Ошибка'
            self.items_name = 'Ошибка артикла'






# -----------------ЗАПУСК
deck = get_sheet()
with open('iter_sort.json', 'r') as file:
    data = json.load(file)

leg = data
# leg = deck.get_all_records()  # Список внутри которого словари (количество словареий равно кличеству строк)
mass = leg[4:]  # Пропускаем пустые строки

deleter = removal(mass)


# Инициализирует таблицу, записывая отсортированные данные в json iter_sort.json
table_data = initor(deleter)  # Получем список словарей (всех строк таблицы)

# Очищаем все строки.
def_table(deck, leg)

# Окрашиваем все строки в белый.
def_color(deck,leg)



def str_generator(table_data):
    '''Данная функция служит фабрикой, оборачивая каждую строку таблицы в классы Strok'''

    class_list = []

    for i in table_data:
        # Класс принимает два аргмента, словрь с данынми о строке, и текущую доску.
        data = Stroka(i, deck)
        class_list.append(data)


    return class_list

def multi_update(obj):
    if bool(obj.name) != False:
        import sys
        obj.product_name_update()



# Фабрика, генератор экземпляров строки.
factory = str_generator(table_data)


with Pool(3) as p:
    p.map(multi_update, factory)


# Кластеризуем экземпляры класса, и обновляем сумму.
klaster = group_data(factory)  # Возвращает список из кластеров (экземпляров класса). [[Вася пупкин],[Вася Петров]]
order_amount(klaster, deck)  # Подсчитывает суммы, вносит изменения в таблицу.

# Обновляем все строки.

def multi(elem):
    elem.table_update()
    time.sleep(1)

with Pool(3) as p:
    p.map(multi, factory)



print(time.time() - start)