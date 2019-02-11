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
    try:
        request = session.get(url, headers=headers, timeout=1)

    except Exception as error:
        time.sleep(1)
        request = session.get(url)

    return request.content


def get_sheet(tokken):
    '''отдаёт экземпляр доски, полученый по токкену'''

    scope = ['https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(tokken, scope)
    client = gspread.authorize(creds)
    sheet = client.open('test1').sheet1
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
        time.sleep(0.2)

        namers = [self.time, self.name, self.doing, self.items_name, self.article, self.item_url, self.unit_item,
                  self.item_value, self.min_value, self.delivery, '', '', self.notice]

        if bool(self.name) != False:

            cell_list = self.sheet.range('A{}:M{}'.format(self.str_number, self.str_number))

            # Обновляем каждоу строку.
            for cell, name in zip(cell_list, namers):
                cell.value = name

            # Update in batch
            self.sheet.update_cells(cell_list)



    def product_name_update(self, sort_dict):
        """Данный метод получает данные о товаре (если он в наличии) обновляет"""
        table_numbers = 'A{}:M{}'.format(self.str_number, self.str_number)
        arrticle = int(self.article)

        if arrticle in sort_dict:
            data = sort_dict.get(arrticle)

            #  Обвноляем цену
            self.unit_item = data.get('price')
            self.items_name = data.get('name')

            # Проверка на товар партнёра
            if data.get('is_remote_store') == 1:
                default_format = CellFormat(backgroundColor=color(250, 10, 10), textFormat=textFormat(bold=False))
                format_cell_range(self.sheet, table_numbers, default_format)
                time.sleep(0.6)

        else:
            #  Если товар НЕ найден в списке артиклов, значит его НЕТ в наличии.
            self.unit_item = 0
            default_format = CellFormat(backgroundColor=color(1, 0, 240), textFormat=textFormat(bold=True))
            format_cell_range(self.sheet, table_numbers, default_format)
            time.sleep(0.6)



        # status = data[4]
        #
        # if status == 'На складе достаточно':
        #     pass





# -----------------ЗАПУСК
# block_access()

login1 = 'login1.json'
login2 = 'login2.json'
login3 = 'login3.json'
login4 = 'login4.json'

deck1 = get_sheet(login1)
deck2 = get_sheet(login2)
deck3 = get_sheet(login3)
deck4 = get_sheet(login4)

leg = deck1.get_all_records()  # Список внутри которого словари (количество словареий равно кличеству строк)
mass = leg[4:]  # Пропускаем пустые строки


deleter = removal(mass)


# Инициализирует таблицу, записывая отсортированные данные в json iter_sort.json
table_data = initor(deleter)  # Получем список словарей (всех строк таблицы)

# Очищаем все строки.
def_table(deck1, leg)

# Окрашиваем все строки в белый.
def_color(deck1,leg)



def str_generator(table_data):
    '''Данная функция служит фабрикой, оборачивая каждую строку таблицы в классы Strok'''

    class_list = []
    vibor = int(0)
    for i in table_data:
        if vibor == 0:
            data = Stroka(i, deck1)
            class_list.append(data)
            vibor += 1
        elif vibor == 1:
            data = Stroka(i, deck2)
            class_list.append(data)
            vibor += 1

        elif vibor == 2:
            data = Stroka(i, deck3)
            class_list.append(data)
            vibor += 1

        else:
            data = Stroka(i, deck4)
            class_list.append(data)
            vibor = 0


    return class_list

# Фабрика, генератор экземпляров строки.
factory = str_generator(table_data)


sorted_dict = get_goods_data(factory)



def multi_update(obj):
    if bool(obj.name) != False:
        obj.product_name_update(sorted_dict)



with Pool(4) as p:
    p.map(multi_update, factory)


# Кластеризуем экземпляры класса, и обновляем сумму.
klaster = group_data(factory)  # Возвращает список из кластеров (экземпляров класса). [[Вася пупкин],[Вася Петров]]


order_amount(klaster)  # Подсчитывает суммы, вносит изменения в таблицу.

# Обновляем все строки.

def multi(elem):
    elem.table_update()


with Pool(4) as p:
    p.map(multi, factory)


# block_access()
print(time.time() - start)