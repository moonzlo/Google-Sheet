from gspread_formatting import *
import html5lib
from operator import itemgetter
import requests, json, time, gspread
from multiprocessing.dummy import Pool
from selenium import webdriver
from requests_xml import XMLSession



def sid_gener(str_list):
    """Генерирует список из артиклов, для передачи в реквест"""
    s = str_list.copy()

    sid_list = []

    try:
        for x in s:
            if bool(x.name) == True:
                sid_list.append(x.article)

        # Удаляем дубликаты артиклов.
        value = set(sid_list)
        value1 = list(value)

        return value1

    except Exception as error:
        print(error)




def sima_api(spisok_sto):
    session = XMLSession()
    r = session.get('https://www.sima-land.ru/api/v3/item/?per-page=?&sid={}'.format(spisok_sto))
    a = r.json()  # Словаря с данными

    return a


def art_sort(artikles):
    """Сортируем артиклы в список по сто штук"""
    schetcik = 0
    spisok = []
    value = []

    for i in artikles:
        if schetcik != 1000:
            spisok.append(i)
            schetcik += 1
        else:
            value.append(','.join(spisok))
            spisok.clear()
            schetcik = 0

    value.append(spisok)

    return value

def get_goods_data(obj_list):
    """Получаем и обрабатываем данные с api сималэнда по списку артиклов, отдаём список товаров в НАЛИЧИИ"""

    stroki = obj_list
    sids = sid_gener(stroki)   # Получаем все артиклы без дублей

    sort_art = art_sort(sids)  # Сортируем артиклы по 1000

    slovar = {}
    for x in sort_art:
        arttiklis = ''.join(str(x).replace(' ', '').replace('[', '').replace(']', ''))
        data = sima_api(arttiklis)
        slovar.update(data)

    sort = []
    for i in slovar.get('items'):
        sort.append(i)

    dict_sort = {}

    for x in sort:
        sid = x.get('sid')
        value = {sid: x}
        dict_sort.update(value)

    return dict_sort


def block_access():  # Блокирует доступ гугл форме.
    """Смысл функции ограничить возмонжость клиентов добавлять контент, во время работы скрипта"""

    options = webdriver.ChromeOptions()
    options.add_argument('headless')  # Без GUI
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument(r"user-data-dir=/home/moonz/git/Google-Sheet/profile/")
    driver = webdriver.Chrome(executable_path="/home/moonz/chromedriver", chrome_options=options)

    # Адресс формы
    driver.get("https://docs.google.com/forms/d/14xEH6imVJgBJKndemdysvTu9olMJEzv0cz62y9ziBdw/edit#responses")



    try:
        driver.find_element_by_xpath(u"(.//*[normalize-space(text()) and normalize-space(.)='Принимать ответы'])[1]/following::div[6]").click()


    except Exception as error:
        print('ОШИБКА ! | ',error)

    finally:
        driver.close()
        driver.quit()


def exchange_rate():
    '''Получает актуальный курс рос.рубля к бел.рублю'''
    html = get_htmls()
    value = []

    doc = html5lib.parse(html, treebuilder="lxml", namespaceHTMLElements=False)
    for node in doc.iterfind("//*span[@class='cc-result']"):
        value.append(node.text)

    kurs = value[1].split()
    return kurs[0]


def final_price(exchange_rate, rub):


    price = float(exchange_rate)*rub

    procent = price / 100 * 8
    total = float(price) + float(procent)

    test = '%.2f' % total
    return test


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

def order_amount(sorted_list):
    '''Данная функция сумирует все цены, и выводит общую сумму заказа'''
    kurs = exchange_rate()


    def core(li):

        SUM = 0
        for i in li:
            deliver = i.delivery
            sheet = i.sheet

            if bool(i.name) != False:

                if type(deliver) == int:

                    vale = str(i.unit_item).strip()
                    amount = i.item_value
                    summin = float(amount) * float(vale)
                    SUM += summin + deliver
                else:
                    vale = str(i.unit_item).strip()
                    amount = i.item_value
                    summin = float(amount) * float(vale)
                    SUM += summin


            else:
                total = final_price(kurs, SUM)
                namers = [SUM, total]
                cell_list = i.sheet.range('K{}:L{}'.format(i.str_number, i.str_number))
                for cell, name in zip(cell_list, namers):
                    cell.value = name
                i.sheet.update_cells(cell_list)
                time.sleep(0.2)
                default_format = CellFormat(backgroundColor=color(30, 10, 10), textFormat=textFormat(bold=True))
                format_cell_range(sheet, 'K{}:L{}'.format(i.str_number, i.str_number), default_format)
                SUM = 0
                time.sleep(0.6)

    with Pool(4) as p:
        p.map(core, sorted_list)


def def_table(deck, table_data):
    '''Цель данной функции, предварительная очистка всего поля от всех данных'''
    stop = len(table_data) + 5
    value = 'A6:M{}'.format(stop)
    cell_list = deck.range(value)

    for cell in cell_list:
        cell.value = ''

    deck.update_cells(cell_list)

    table_numbers = 'A6:M{}'.format(stop)
    default_format = CellFormat(backgroundColor=color(1, 1, 1), textFormat=textFormat(bold=False))
    format_cell_range(deck, table_numbers, default_format)


def initor(sheet):
    """Сначала удаляем те данные которые Хотят удалить из таблицы, сми клиенты"""
    data = sheet
    black = []

    for i in data:
        if i.get('Дейсвтие') == 'Удаляю':
            value = []
            value.append(i)
            data.remove(i)
            for x in data:
                if value[0].get('Фамилия и имя') == x.get('Фамилия и имя') and value[0].get('Артикул') == x.get('Артикул'):
                    data.remove(x)

    '''сортирует все данные по алфовиту + добавляет пустые строки в конеце кластера'''

    # Получаем отсортированный по именам список. Пустые поля попадают в самый верх.
    list_of_dicts = data
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
    # with open('iter_sort.json', 'w') as file:
    #     json.dump(sorted_list, file, indent=2, ensure_ascii=False)

    return sorted_list