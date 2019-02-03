from gspread_formatting import *
from bs4 import BeautifulSoup
from operator import itemgetter
import requests, json, time


def exchange_rate():
    '''Получает актуальный курс рос.рубля к бел.рублю'''
    html = get_htmls()
    soup = BeautifulSoup(html, 'html.parser')

    tag = soup.find_all('span', class_='cc-result')
    a = tag[1].text.strip()

    b = re.findall(r'[0-9.]', '{}'.format(a))
    num = str().join(b)
    number = float(num)

    return number



def final_price(exchange_rate, rub):

    price = exchange_rate*rub

    procent = price / 100 * 8
    total = int(price) + int(procent)
    return total


def get_htmls():
    url = 'https://rub.ru.currencyrate.today/byn'

    headers = {'accept': '*/*',
               'user-agent':
                   'Mozilla/5.0 (Windows NT 6.0) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.52 Safari/536.5'}
    session = requests.Session()  # Иметирует сессию клиента.
    request = session.get(url, headers=headers)

    return request.content



# Функция принимает на вход список из экземпляров класса Stroka

def group_data(factory_list):
    '''Цель данной функции, объеденить строки в списки, для удобства взаимодействия и расчётов'''
    sorted_list = []  # Список кластеризованных экземпляров класса Stroka, в списки.

    klaster = []

    for x in factory_list:
        if bool(x.name) != False:
            klaster.append(x)
        else:
            klaster.append(x)
            sorted_list.append(klaster)
            klaster = []

    return sorted_list


# Запускать в самом конце.

def order_amount(sorted_list, sheet):
    '''Данная функция сумирует все цены, и выводит общую сумму заказа'''

    summ = []
    for li in sorted_list:
        for i in li:
            if bool(i.name) != False:

                vale = str(i.unit_item).strip()
                amount = i.item_value
                summin = int(amount) * int(vale)
                summ.append(summin)

            else:
                # Обновляем итог, красим ячейку в дефолтный цвет и выделяем жиром.
                sheet.update_acell('K{}'.format(i.str_number), sum(summ))
                default_format = CellFormat(backgroundColor=color(30, 10, 10), textFormat=textFormat(bold=True))
                format_cell_range(sheet, 'K{}'.format(i.str_number), default_format)

                # Обвноляем данные орг.сбора и конвертация в бел.рубль
                total = final_price(exchange_rate(), sum(summ))

                sheet.update_acell('L{}'.format(i.str_number), total)
                default_format = CellFormat(backgroundColor=color(30, 10, 10), textFormat=textFormat(bold=True))
                format_cell_range(sheet, 'L{}'.format(i.str_number), default_format)


def removal(table_data):
    data = table_data
    black = []

    for i in data:
        if i.get('Дейсвтие') == 'Удаляю':
            value = []
            value.append(i)
            data.remove(i)
            for x in data:
                if value[0].get('Фамилия и имя') == x.get('Фамилия и имя') and value[0].get('Артикул') == x.get('Артикул'):
                    data.remove(x)


    return data

def def_color(deck, table_data):
    '''Данная функция сулжит для предварительного окраса всех строк в белый цвет'''
    num = len(table_data)
    table_numbers = 'A6:M{}'.format(num)
    default_format = CellFormat(backgroundColor=color(1, 1, 1), textFormat=textFormat(bold=False))
    format_cell_range(deck, table_numbers, default_format)

def def_table(deck, table_data):
    '''Цель данной функции, предварительная очистка всего поля от всех данных'''
    stop = len(table_data) + 1
    value = 'A6:M{}'.format(stop)
    cell_list = deck.range(value)

    for cell in cell_list:
        cell.value = ' '

    deck.update_cells(cell_list)


def initor(sheet):
    '''сортирует все данные по алфовиту + добавляет пустые строки в конеце кластера'''

    # Получаем отсортированный по именам список. Пустые поля попадают в самый верх.
    list_of_dicts = sheet
    # list_of_dicts.sort(key=itemgetter('Фамилия и имя'))
    # Убираем все пустые строки, и сортируем по алфовиту.
    new = []
    new.extend(list_of_dicts)
    sorted_list_new = []

    for i, c in enumerate(new):
        a = str(c.get('Фамилия и имя')).strip()

        if bool(a) == True:
            sorted_list_new.append(c)


    sorted_list_new.sort(key=itemgetter('Фамилия и имя'))


    def work_list(list_dict):
        # Пустая строка разделитель.
        zaglushka ={
            "Отметка времени": "",
            "Фамилия и имя": "",
            "Действие": " ",
            "Наименование (кратко)": " ",
            "Артикул": " ",
            "Ссылка на товар": " ",
            "Цена за единицу товара": " ",
            "Количество (нужное Вам)": " ",
            "Количество для минимального заказа": " ",
            "Примечание": "",
            "Цена за доставку (если есть)": " ",
            "Сумма заказа в РОС.РУБЛЯХ": " ",
            "Итоговая сумма (+орг.сбор 9%)": " "
        }

        sorted_dict_list = []

        for i, c in enumerate(list_dict):
            try:

                if list_dict[i].get('Фамилия и имя') != list_dict[i + 1].get('Фамилия и имя'):
                    sorted_dict_list.append(c)
                    a = zaglushka.copy()
                    a.update({'Заглушка для:':list_dict[i].get('Фамилия и имя')})
                    sorted_dict_list.append(a)

                else:
                    sorted_dict_list.append(c)
            except IndexError:  # Когда кончается список
                sorted_dict_list.append(c)
                a = zaglushka.copy()
                a.update({'Заглушка для:': list_dict[i].get('Фамилия и имя')})
                sorted_dict_list.append(a)


        return sorted_dict_list  # Возвращаем отсортированные список словарей + загрулшки.

    # Передаем в функцию отсортированные по алфовиту, получаем записанные файл и список с загрушками.
    sorted_list = work_list(sorted_list_new)

    # Добавляем индекс строки в конце каждого словаря.

    nums = 6  # Начальная строка, от которой будет идти обновление
    for i, c in enumerate(sorted_list):
        sorted_list[i].update({'Строка':nums})
        nums += 1

    # Записываем пронумерованные строки в файл.
    with open('iter_sort.json', 'w') as file:
        json.dump(sorted_list, file, indent=2, ensure_ascii=False)

    return sorted_list