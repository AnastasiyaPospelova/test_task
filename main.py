import pyexcel as pe
from datetime import datetime
from pyexcel_xls import save_data
from collections import OrderedDict


clients_table = pe.get_sheet(file_name="Тест_Клиенты текущие.ods")  # Загружаем данные из таблицы Клиенты

flag = False
deals = []
for index, lst in enumerate(clients_table):  # Выбираем из таблицы Клиенты список сделок второй группы
    if "Группа номер два" in lst:
        flag = True
    if flag and lst[1] != '':
        if lst[2] > datetime.strptime('2024-06-01', '%Y-%m-%d'):
            deals.append([lst[1], lst[5], lst[3]])
# print("Задание 1: выбор сделок группы номер два.\n", deals)


results_table = pe.get_sheet(file_name="Тест_Итог.ods")  # Загружаем данные из таблицы Итог


def check_client(table_deals: list, deal_number) -> bool:  # Проверка отсутствия клиента в таблице
    for j in table_deals:
        if j == '':
            return True
        elif j == deal_number:
            return False


flag = False
index = 0
for i, lst in enumerate(results_table):  # Добавляем в табл. Итог информацию о сделках из табл. Клиенты
    if lst[1] == '':
        flag = True

    if flag and index != len(deals):
        while not check_client(list(results_table.columns())[2], deals[index][1]):
            index += 1
        results_table[i, 1] = deals[index][0]
        results_table[i, 2] = deals[index][1]
        results_table[i, 3] = deals[index][2]
        index += 1

# print("\nНовая таблица клиентов:\n", results_table)
results_dict = OrderedDict()
results_dict.update({"Сделки": list(results_table.rows())})
save_data("Тест_Итог.xls", results_dict)  # Сохранение обновлённой таблицы Итог


def get_actual_clients(table: pe.sheet.Sheet) -> list:  # Получение списка актуальных клиентов
    actuals = []
    current_datetime = datetime.now()
    for i, lst in enumerate(table):
        if not i or lst[0] == '':
            pass
        elif lst[1] > current_datetime:
            actuals.append(lst[0])
    return actuals


actuality_table = pe.get_sheet(file_name="Тест_Актуальность.ods")  # Загружаем данные из таблицы Актуальность
numbers_clients = clients_table.column_at(1)
branches_dict = {}
actual_clients = get_actual_clients(actuality_table)

for i in numbers_clients:  # Формирование сумм сделок по филиалам
    if str(i).isdigit():
        if i in actual_clients:
            index = numbers_clients.index(i)
            if clients_table[index, 4] not in branches_dict.keys():
                branches_dict[clients_table[index, 4]] = clients_table[index, 3]
            else:
                branches_dict[clients_table[index, 4]] += clients_table[index, 3]

branches_list = [[key, val] for key, val in branches_dict.items()]
branches_list.insert(0, ['Филиал', 'Сумма'])
clients_table_list = list(clients_table.rows())
clients_table_list[0].insert(1, '')

results_dict = OrderedDict()
results_dict.update({"Клиенты": list(clients_table.rows())})
results_dict.update({"Филиалы": branches_list})
save_data("Тест_Клиенты текущие.xls", results_dict)  # Сохранение обновлённой книги Клиенты
