import xlrd
import pandas as pd
import re
import datetime

c_shops = {'США': ['6РМ', 'AMAZON', 'МК', 'СК', 'ASHFORD', 'AX', 'BF', 'BR', 'CARTERS', 'CHP', 'CK', 'COACH', 'DISNEY',
                   'DKNY', 'DOPESNOW', 'GAP_OLD', 'GF', 'GF-BF', 'GF-BR', 'GYM', 'JJ', 'JOULES', 'KATESPADE', 'KL',
                   'LEVIS', 'LD', 'MACYS', 'MK', 'NB', 'NORD', 'OLD', 'OLD-BANANA', 'OLD-GAP', 'OLD(GAP) ', 'RALPH',
                   'RL', 'SAKS', 'SHOPDIABETS', 'TIMB', 'TOMMY', 'USPOLO', 'VIVAIA', 'VS'],
           'Германия': ['СА', 'FAMILY', 'AVARCA', 'BS', 'CA', 'DM', 'PINKO', 'PROSHOP', 'SOFTSHELL', 'WORLDOFSWEETS',
                        'ZALANDO'],
           'Англия': ['AMAZON_UK', 'ASDA', 'M&S', 'MATALAN', 'TU'],
           'Португалия': ['CORTEFIEL', 'DESIGUAL', 'MASSIMO', 'WS'],
           'Россия': ['Авторские адвент-календари', 'Авторский дизайн от Captains Mate', 'Арахисовые пасты',
                      'Ароматный кофе', 'Банты ручной работы', 'Белевские сладости', 'Бомбочки для ванны',
                      'Восточные сладости', 'Домашняя одежда', 'Есть в наличии', 'Настя и Никита', 'Издательство МИФ',
                      'Нигма', 'Стрекоза', 'ARCA', 'Мелик-Пашаев', 'Пешком в историю', 'Поляндрия',
                      'Корейская косметика', 'Косметика Корея', 'Косметика Green Era', 'G.LOVE', 'SmoRodina',
                      'Сокровища пиратов', 'Необычный чай', 'Пазлы DaVici', 'Uzelki', 'Шанти Пунти', 'Polezzno',
                      'Сарапульский кондитер', 'Laboratorium', 'Сибирский кедр', 'TRAWA', 'Термосы KLEAN KANTEEN',
                      'Huggy', 'Шанти-Пунти', 'AIM', 'Crockid', 'Estel', 'LOCAL']}

month_dict = {'января': '01', 'февраля': '02', 'марта': '03', 'апреля': '04', 'мая': '05', 'июня': '06',
              'июля': '7', 'августа': '8', 'сентября': '9', 'октября': '10', 'ноября': '11', 'декабря': '12',
              'дек': '12'}

sv = {'MK': 'МК', 'CK': 'СК', 'CA': 'СА'}
temp = dict()
for k, v in sv.items():
    temp[v] = k
sv = sv | temp

shop_cs = {}
for k, v in c_shops.items():
    for x in v:
        shop_cs[x] = k


book = xlrd.open_workbook("D:\\job\\hidream\\var2\\result.xls")
# print("The number of worksheets is {0}".format(book.nsheets))
# print("Worksheet name(s): {0}".format(book.sheet_names()))
var2 = book.sheet_by_index(0)
data_for_var2 = book.sheet_by_index(1)
# print("{0} {1} {2}".format(var2.name, var2.nrows, var2.ncols))
# print("Cell A2 is {0}".format(var2.cell_value(rowx=3, colx=0)))
# print("{0} {1} {2}".format(data_for_var2.name, data_for_var2.nrows, data_for_var2.ncols))
# print("Cell A2 is {0}".format(data_for_var2.cell_value(rowx=3, colx=0)))

output = pd.DataFrame(columns=['ID', 'Название', 'Направление', 'Контакт', 'Товар', 'Цена за 1 шт', 'Количество',
                               'Агентское вознаграждение', 'Итого', 'Стадия сделки', 'Код посылки наш',
                               'Суммарный вес', 'Страна', 'Код посылки склада', 'Дата аванса', 'Сумма аванса',
                               'Выдали покупку', 'Код покупки'])

data_errs = []


def find_code(name, product, price, quantity):
    global data_for_var2
    global data_errs
    name_two = ' '.join(name.split()[::-1])
    for row in range(data_for_var2.nrows):
        if ((name.lower() in data_for_var2.cell_value(rowx=row, colx=2).lower()
             or name_two.lower() in data_for_var2.cell_value(rowx=row, colx=2).lower())
            and data_for_var2.cell_value(rowx=row, colx=3) == product
            and data_for_var2.cell_value(rowx=row, colx=4) == price
            and data_for_var2.cell_value(rowx=row, colx=5) == quantity):
            country = data_for_var2.cell_value(rowx=row, colx=11)
            avsum = data_for_var2.cell_value(rowx=row, colx=14)
            code = data_for_var2.cell_value(rowx=row, colx=18)
            avdate = data_for_var2.cell_value(rowx=row, colx=13)
            avdate = datetime.datetime(*xlrd.xldate_as_tuple(avdate, book.datemode))
            avdate = avdate.strftime("%d.%m.%y")
            return country, avdate, avsum, code
    data_errs.append([name, product, price, quantity])
    return None


deal_id = 135
direction = 'Общее'
contact = ''
product_name = ''
price = 0
quantity = 1
agent = 0
total = 0
status = 'Сделка успешна'
our_pc = 1
total_weight = 1
country = 'None'
wh_pc = 1
avdate = ''
avsum = 0
date = ''
code = ''
deal_name = ''

date_q = False
good_time = False
prod_id = 0

for row in range(1, var2.nrows):
    prod_id = int(var2.cell_value(rowx=row, colx=0))
    deal_id = var2.cell_value(rowx=row, colx=1)
    deal_name = var2.cell_value(rowx=row, colx=2)
    direction = var2.cell_value(rowx=row, colx=3)
    contact = var2.cell_value(rowx=row, colx=4)
    product_name = var2.cell_value(rowx=row, colx=5)
    price = var2.cell_value(rowx=row, colx=6)
    quantity = var2.cell_value(rowx=row, colx=7)
    agent = var2.cell_value(rowx=row, colx=8)
    total = var2.cell_value(rowx=row, colx=9)
    status = var2.cell_value(rowx=row, colx=10)
    our_pc = var2.cell_value(rowx=row, colx=11)
    total_weight = var2.cell_value(rowx=row, colx=12)
    country = var2.cell_value(rowx=row, colx=13)
    wh_pc = var2.cell_value(rowx=row, colx=14)
    avdate = var2.cell_value(rowx=row, colx=15)
    avsum = var2.cell_value(rowx=row, colx=16)
    date = var2.cell_value(rowx=row, colx=17)
    code = var2.cell_value(rowx=row, colx=18)
    if country == '' or avdate == '' or avsum == '' or code == '':
        if find_code(contact, product_name, price, quantity) is not None:
            tmp_country, tmp_avdate, tmp_avsum, tmp_code = find_code(contact, product_name, price, quantity)
            if country == '':
                country = tmp_country
            if avdate == '':
                avdate = tmp_avdate
            if avsum == '':
                avsum = tmp_avsum
            if code == '':
                code = tmp_code

    output.loc[prod_id] = [deal_id, deal_name, direction, contact, product_name, price, quantity, agent, total,
                           status, our_pc, total_weight, country, wh_pc, avdate, avsum, date, code]


print('Ошибки получения данных')
output2 = pd.DataFrame(columns=['Контакт', 'Товар', 'Цена за 1 шт', 'Количество'])
print(data_errs)
idd = 0
for i in data_errs:
    output2.loc[idd] = [i[0], i[1], i[2], i[3]]
    idd += 1
# print('\n\n\n')
# output.to_excel('./result2.xlsx')
output2.to_excel('./errors.xlsx')
