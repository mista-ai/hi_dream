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


def find_date(date_cell):
    words = ''
    try:
        words = re.split('\W+', date_cell.lower())
    except:
        print('Hello error')
    global month_dict
    months = month_dict.keys()
    mindex = -1
    for month in months:
        if month in words:
            mindex = words.index(month)
            break
    if mindex <= 0:
        raise ValueError
    result = words[mindex - 1].rjust(2, '0') + '.' + month_dict[words[mindex]].rjust(2, '0') + '.21'
    return result


book = xlrd.open_workbook("D:\\job\\hidream\\TZ.xls")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(2)
rs = book.sheet_by_index(3)
report = book.sheet_by_index(1)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
print("Cell A2 is {0}".format(sh.cell_value(rowx=3, colx=0)))

code_errs = []
avdate_errs = []


def find_code(name, product):
    global report
    global code_errs
    global c_shops
    global shop_cs
    name_two = ' '.join(name.split()[::-1])
    for row in range(report.nrows):
        if ((name.lower() in report.cell_value(rowx=row, colx=1).lower()
             or name_two.lower() in report.cell_value(rowx=row, colx=1).lower())
            and report.cell_value(rowx=row, colx=3) == product):
            code = report.cell_value(rowx=row, colx=0)
            if 'доставка' in code.lower():
                for country in c_shops.keys():
                    if country.lower() in code.lower():
                        return [code, country]
            elif 'отправка' in code.lower():
                return [code, 'Россия']
            else:
                if product == 'Детская бутылка Klean Kanteen Kid Classic Sippy 12oz (355 мл) Millennial Hearts':
                    print('Hi man')
                for shop in shop_cs.keys():
                    if shop.lower() in code.lower():
                        return [code, shop_cs[shop]]
    code_errs.append([name, product])
    return None, None


def find_avdate(name, code):
    global rs
    global avdate_errs
    name_two = ' '.join(name.split()[::-1])
    for row in range(rs.nrows):
        if code is not None:
            for key in sv.keys():
                if key in code:
                    temp = code.replace(key, sv[key])
                    if ((name.lower() in rs.cell_value(rowx=row, colx=4).lower()
                         or name_two.lower() in rs.cell_value(rowx=row, colx=4).lower())
                        and (code.lower() in rs.cell_value(rowx=row, colx=5).lower()
                             or temp.lower() in rs.cell_value(rowx=row, colx=5).lower())):
                        # дата аванса, сумма аванса
                        dater = rs.cell_value(rowx=row, colx=0)
                        dater = datetime.datetime(*xlrd.xldate_as_tuple(dater, book.datemode))
                        dater = dater.strftime("%d.%m.%y")
                        avsum = rs.cell_value(rowx=row, colx=1)
                        print(dater, avsum)
                        return dater, avsum
                    break

            if ((name.lower() in rs.cell_value(rowx=row, colx=4).lower()
                 or name_two.lower() in rs.cell_value(rowx=row, colx=4).lower())
                and (code.lower() in rs.cell_value(rowx=row, colx=5).lower())):
                # дата аванса, сумма аванса
                dater = rs.cell_value(rowx=row, colx=0)
                dater = datetime.datetime(*xlrd.xldate_as_tuple(dater, book.datemode))
                dater = dater.strftime("%d.%m.%y")
                avsum = rs.cell_value(rowx=row, colx=1)
                print(dater, avsum)
                return dater, avsum

    avdate_errs.append([name, code])
    return None, None


output = pd.DataFrame(columns=['ID', 'Название', 'Направление', 'Контакт', 'Товар', 'Цена за 1 шт', 'Количество',
                               'Агентское вознаграждение', 'Итого', 'Стадия сделки', 'Код посылки наш',
                               'Суммарный вес', 'Страна', 'Код посылки склада', 'Дата аванса', 'Сумма аванса',
                               'Выдали покупку', 'Код покупки'])

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

for row in range(sh.nrows):
    cell = sh.cell_value(rowx=row, colx=0)
    if type(cell) is str:
        if len(cell.split(' ')) == 2 and sh.cell_value(rowx=row, colx=1) == '':
            deal_id += 1
            contact = cell
            deal_name = 'Сделка #' + str(deal_id)
            date_q = True
            good_time = False
            continue
    if date_q:
        try:
            date = find_date(sh.cell_value(rowx=row, colx=0))
            # date = datetime.datetime(*xlrd.xldate_as_tuple(date, book.datemode))
            # date = date.strftime("%d.%m.%y")
        except:
            print(sh.cell_value(rowx=row, colx=0), row)
        date_q = False
        good_time = True
        continue

    if good_time:
        if cell == '' or sh.cell_value(rowx=row, colx=1) == '':
            continue
        elif 'пожалуйста' in cell.lower():
            good_time = False
        elif 'доставка' in cell.lower() or 'отправка' in cell.lower():
            product_name = cell
            price = sh.cell_value(rowx=row, colx=1)
            quantity = 1
            agent = 0
            try:
                total = float(price) * float(quantity)
            except:
                print(product_name, price, 'row=', row, ' |', quantity)
            code, country = find_code(contact, product_name)
            avdate, avsum = find_avdate(contact, code)
            output.loc[prod_id] = [deal_id, deal_name, direction, contact, product_name, price, quantity, agent, total,
                                   status, our_pc, total_weight, country, wh_pc, avdate, avsum, date, code]
            prod_id += 1

        else:
            product_name = cell
            price = sh.cell_value(rowx=row, colx=1)
            quantity = sh.cell_value(rowx=row, colx=3)
            agent = sh.cell_value(rowx=row, colx=6)
            try:
                total = float(price) * float(quantity)
            except:
                print(product_name, price, quantity, 'row=', row, 'доставка' in cell.lower())
            code, country = find_code(contact, product_name)
            avdate, avsum = find_avdate(contact, code)
            # df = pd.DataFrame({'ID': deal_id, 'Направление': direction, 'Контакт': contact, 'Товар': product_name,
            #                    'Цена за 1 шт': price, 'Количество': quantity, 'Агентское вознаграждение': agent,
            #                    'Итого': total, 'Стадия сделки': status, 'Код посылки наш': our_pc,
            #                    'Суммарный вес': total_weight, 'Страна': country, 'Код посылки склада': wh_pc,
            #                    'Дата аванса': avdate, 'Сумма аванса': avsum, 'Выдали покупку': date,
            #                    'Код покупки': code})
            output.loc[prod_id] = [deal_id, deal_name, direction, contact, product_name, price, quantity, agent, total,
                                   status, our_pc, total_weight, country, wh_pc, avdate, avsum, date, code]
            prod_id += 1

# print(pd)
print('Ошибки получения кода')
print(code_errs)
print('\n\n\n')
print('Ошибки получения аванса')
print(avdate_errs)
output.to_excel('./result.xlsx')
