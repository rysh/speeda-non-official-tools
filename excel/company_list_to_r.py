#-*- coding: utf-8 -*-

import openpyxl as px
import csv


def create_year_range():
    return list(range(1970, 2016))


def target_range(sheet_name):
    if sheet_name == '上場企業':
        return 'A11:GI3656'
    elif sheet_name == '上場配当総額':
        return 'A11:BA3656'
    elif sheet_name == '非上場統合':
        return 'A11:GI1258'
    elif sheet_name == '非上場配当総額統合':
        return 'A11:GI1258'
    elif sheet_name == '企業リスト':
        return 'A11:L181'


def position(sheet_name, index):
    if sheet_name == '上場企業' or sheet_name == '非上場統合':
        if index < 53:
            return 6
        elif index < 99:
            return 7
        elif index < 145:
            return 8
        elif index < 191:
            return 9
    elif sheet_name == '上場配当総額' or sheet_name == '非上場配当総額統合':
        return 10


def del_none(val):
    if val is None:
        return ''
    return val


def label_col(sheet_name):
    if sheet_name == '企業リスト':
        return [0, 1, 3, 5, 6]
    return [0, 1, 2, 4, 6]


def write_csv(records):
    with open('results.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile, delimiter=',',
                                quotechar='|', quoting=csv.QUOTE_MINIMAL)
        writer.writerow(
            ['SPEEDA企業ID', 'コード', '企業名称', '業種', '優先市場', '年度', '有利子負債残高', '株主資本等合計', '資産合計', '一株当たり年間配当金', '配当総額'])
        for record in records:
            # print(record)
            writer.writerow(record)


def init_record(labels, year):
    record = [''] * 11
    for label_i, v in enumerate(labels):
        record[label_i] = del_none(v)
    if '銀行・証券・保険' in record[3]:
        record[3] = '金融業'
    else:
        record[3] = ''
    record[5] = year
    return record


def process_worksheet(sheet_name, ws, records):
    for row in ws[target_range(sheet_name)]:

        labels = []

        year_range = create_year_range()
        for i, cell in enumerate(row):
            if i in label_col(sheet_name):
                labels.append(cell.value)

            if sheet_name == '企業リスト':
                if i == 6:
                    record = init_record(labels, '')
                    record[4] = '非上場'
                    records[labels[0]] = record
                continue

            if i > 6:
                if len(year_range) == 0:
                    year_range = create_year_range()
                year = str(year_range.pop(0))
                key = labels[0] + year
                if records.get(key) is None:
                    records[key] = init_record(labels, year)
                records[key][position(sheet_name, i)] = del_none(str(cell.value))

    return records

if __name__ == "__main__":
    files = {
             '上場全社': ['上場企業','上場配当総額'],
             '非上場': ['非上場統合'],
             '非上場配当総額': ['非上場配当総額統合'],
             '非上場（金融）': ['企業リスト'],
             }

    records = {}
    for file_name in files.keys():
        wb = px.load_workbook('./' + file_name + '.xlsx')
        for sheet_name in files[file_name]:
            print('processing ' + file_name + '.' + sheet_name + '...')
            ws = wb.get_sheet_by_name(sheet_name)
            records = process_worksheet(sheet_name, ws, records)

    write_csv(records.values())

