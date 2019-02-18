#!/venv/bin/ python

from oauth2client.service_account import ServiceAccountCredentials
from func import *
from multiprocessing.dummy import Pool
import gc


def get_sheet(tokken, sheet_name, table_name):
    """отдаёт экземпляр доски, полученый по токкену"""

    scope = ['https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(tokken, scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).worksheet(table_name)
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
        self.status = dictionary.get('Состояние')

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



    def product_name_update(self, sort_dict, factory):
        try:
            """Данный метод получает данные о товаре (если он в наличии) обновляет"""
            table_numbers = 'A{}:M{}'.format(self.str_number, self.str_number)
            arrticle = int(self.article)

            if arrticle in sort_dict:
                data = sort_dict.get(arrticle)

                # Правильная ссылка на товар
                self.item_url = f'https://www.sima-land.ru/{self.article}'

                #  Обвноляем цену, имя, и минимальный выкуп
                self.unit_item = data.get('price')
                self.items_name = data.get('name')
                self.min_value = data.get('qty_rules')

                # Проверка цены за доставку, если НЕ указана то проверям статус.
                if bool(self.delivery) == False:
                    deli_status = data.get('is_free_delivery')
                    if deli_status == True:
                        self.delivery = 'Бесплатная'
                    else:
                        self.delivery = 'Платная'

                # Проверка на выкуп
                def sverka(article, min_value, item_value):
                    artikl = article
                    min = min_value.split()

                    all_art = 0
                    all_art += item_value

                    for i in factory:
                        if i.article == artikl:  # Ищем такой же артикл в таблице.
                            # Если нашли совпадение по артиклу, добавим столько сколько хочет купить человек
                            all_art += int(i.item_value)

                    if all_art >= int(min[1]):
                        return True

                    elif int(min[1]) == 1:
                        return True

                    else:
                        return False

                # Проверка, достаточно ли товара для выкупа.
                status = sverka(self.article, self.min_value, self.item_value)
                try:
                    if bool(self.status) == True:
                        if int(self.status) == 1:
                            #  Окращивает линию в зеленый
                            time.sleep(0.8)
                            default_format = CellFormat(backgroundColor=color(250, 10, 0), textFormat=textFormat(bold=False))
                            format_cell_range(self.sheet, table_numbers, default_format)
                        elif int(self.status) == 2:

                            time.sleep(0.8)
                            default_format = CellFormat(backgroundColor=color(100,1,250), textFormat=textFormat(bold=False))
                            format_cell_range(self.sheet, table_numbers, default_format)


                    elif status == False:
                        # Если НЕ хватает для выкупа.
                        time.sleep(0.8)
                        default_format = CellFormat(backgroundColor=color(1, 2, 0), textFormat=textFormat(bold=False))
                        format_cell_range(self.sheet, table_numbers, default_format)

                    # Проверка на товар партнёра
                    elif data.get('is_remote_store') == 1:
                        time.sleep(0.8)
                        default_format = CellFormat(backgroundColor=color(250, 10, 10), textFormat=textFormat(bold=False))
                        format_cell_range(self.sheet, table_numbers, default_format)

                except Exception as error:
                    print(error)

            else:
                #  Если товар НЕ найден в списке артиклов, значит его НЕТ в наличии.
                self.unit_item = 0
                time.sleep(0.9)
                default_format = CellFormat(backgroundColor=color(10, 0, 0), textFormat=textFormat(bold=True))
                format_cell_range(self.sheet, table_numbers, default_format)

        except Exception as error:
            print('Я строка', self.str_number)
            print('Случилась ошибка =( :', error)






def main():
    start = time.time()

    form_url = '14xEH6imVJgBJKndemdysvTu9olMJEzv0cz62y9ziBdw'
    chrome_profile = '/home/moonzlo/PycharmProjects/Google-Sheet/'
    webriver_route = '/home/moonzlo/PycharmProjects/Google-Sheet/chromedriver'
    tokken_route = '/home/moonzlo/PycharmProjects/Google-Sheet/'
    sheet_name = 'test1'  # Имя самой таблицы
    table_name = 'test'  # Имя таблицы внутри

    # Блокируем возможность добавлять данные в таблицу
    block_access(webriver_route, chrome_profile, form_url)

    logins = [tokken_route + 'login1.json', tokken_route + 'login2.json', tokken_route + 'login3.json',
              tokken_route + 'login4.json', ]

    deck1 = get_sheet(logins[0], sheet_name, table_name)
    deck2 = get_sheet(logins[1], sheet_name, table_name)
    deck3 = get_sheet(logins[2], sheet_name, table_name)
    deck4 = get_sheet(logins[3], sheet_name, table_name)

    leg = deck1.get_all_records()  # Список внутри которого словари (количество словареий равно кличеству строк)
    mass = leg[4:]  # Пропускаем пустые строки

    # Инициализирует таблицу, записывая отсортированные данные в json iter_sort.json
    table_data = initor(mass)  # Получем список словарей (всех строк таблицы)

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

    # Очищаем все строки.
    def_table(deck1, factory)

    sorted_dict = get_goods_data(factory)

    def multi_update(obj):

        if bool(obj.name) != False:
            obj.product_name_update(sorted_dict, factory)

    with Pool(4) as p:
        p.map(multi_update, factory)

    # Кластеризуем экземпляры класса, и обновляем сумму.
    klaster = group_data(factory)  # Возвращает список из кластеров (экземпляров класса). [[Вася пупкин],[Вася Петров]]

    def table_update(table, spisok):
        try:
            stop = len(spisok) + 5
            value = f'A6:N{stop}'
            cell_list = table.range(value)
            num1 = 0
            num2 = 0

            for cell in cell_list:
                stroka = spisok[num1]
                namers = [stroka.time, stroka.name, stroka.doing, stroka.items_name, stroka.article, stroka.item_url,
                          stroka.unit_item, stroka.item_value, stroka.min_value, stroka.delivery, '', '', stroka.notice,
                          stroka.status]
                namers2 = ['', '', '', '', '', '', '', '', '', '', '', '', '', '']

                if bool(namers[1]) == True:

                    if num2 != 13:
                        cell.value = namers[num2]
                        num2 += 1
                    else:
                        num2 = 0
                        num1 += 1
                        cell.value = namers[13]

                else:
                    if num2 != 13:
                        cell.value = namers2[num2]
                        num2 += 1
                    else:
                        num2 = 0
                        num1 += 1
                        cell.value = namers2[13]

            table.update_cells(cell_list)

        except Exception as error:
            print(error)

    table_update(deck1, factory)

    order_amount(klaster)  # Подсчитывает суммы, вносит изменения в таблицу.

    # Разблокируем возможность добавлять данные в таблицу
    block_access(webriver_route, chrome_profile, form_url)
    return time.time() - start


if __name__ == "__main__":
    try:
        print(main())
    except Exception as error:
        print(error)
        # import requests
        # get = requests.get(f'https://api.telegram.org/bot718325311:AAGQ0ixXKaV9lKGJZLHWr5eAhKrL1gOpeCc/sendMessage?chat_id=191494526&text={error}')


