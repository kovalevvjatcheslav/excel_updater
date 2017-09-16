# -*- coding: utf8 -*-

import pywintypes
import win32com.client
import os.path
import re

excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
path_to_scan = u'\\\\files.skb-turbina.com\shares\ОГК\\43-й отдел\Внутренняя переписка СКАН'

# паттерн поиска записи документа и файла документа
re_pattern = re.compile(u'43[-.][0-9]+[-.][0-9]{2,4}')

# паттерны поиска записи о дополнениях или файла дополнения
re_pattern1 = re.compile(u'Доп\. №([0-9]).*(43[-.][0-9]+[-.][0-9]{2,4})')
re_pattern2 = re.compile(u'Д([0-9]) (43[-.][0-9]+[-.][0-9]{2,4})')
re_pattern3 = re.compile(u'(43[-.][0-9]+[-.][0-9]{2,4}) Д([0-9])')

log_file = open('log', 'w')
try:
    work_book = excel.Workbooks.Open(os.path.join(u'\\\\files.skb-turbina.com\shares\ОГК\\43-й отдел', u'Корреспонденция ОТД. №43.xlsx'))
    work_sheet = work_book.Worksheets(u'С.З., Акты, Решения, О.З., Отч.')
    row = 2
    while True:
        cell = 'B{0}'.format(row)
        value = work_sheet.Range(cell).Value
        if not value:
            break
        try:
            hyperlink = work_sheet.Range(cell).Hyperlinks(1).Address
        except pywintypes.com_error as e:
            # TODO: вынести повторяющийся код в отдельную функцию
            if re.findall(re_pattern1, value):
                additional_count = re.findall(re_pattern1, value)[0][0]
                doc_number = re.findall(re_pattern1, value)[0][1]
                for path_to_file in os.listdir(path_to_scan):
                    if re.findall(re.compile(u'Доп\. №{0}(.)*{1}'.format(additional_count, doc_number)), path_to_file):
                        work_sheet.Hyperlinks.Add(Anchor=work_sheet.Range(cell),
                                                  Address=os.path.join(path_to_scan, path_to_file))
                        work_sheet.Range(cell).Font.Name = 'Times New Roman'
                        work_sheet.Range(cell).Font.Size = 14
                        work_sheet.Range(cell).HorizontalAlignment = win32com.client.constants.xlCenter
                        work_sheet.Range(cell).VerticalAlignment = win32com.client.constants.xlCenter
                        log_file.write(u'{0}\n'.format(value).encode('cp1251'))
                    elif re.findall(re.compile(u'Д{0} {1}'.format(additional_count, doc_number)), path_to_file):
                        work_sheet.Hyperlinks.Add(Anchor=work_sheet.Range(cell),
                                                  Address=os.path.join(path_to_scan, path_to_file))
                        work_sheet.Range(cell).Font.Name = 'Times New Roman'
                        work_sheet.Range(cell).Font.Size = 14
                        work_sheet.Range(cell).HorizontalAlignment = win32com.client.constants.xlCenter
                        work_sheet.Range(cell).VerticalAlignment = win32com.client.constants.xlCenter
                        log_file.write(u'{0}\n'.format(value).encode('cp1251'))
                    elif re.findall(re.compile(u'{0} Д{1}'.format(doc_number, additional_count)), path_to_file):
                        work_sheet.Hyperlinks.Add(Anchor=work_sheet.Range(cell),
                                                  Address=os.path.join(path_to_scan, path_to_file))
                        work_sheet.Range(cell).Font.Name = 'Times New Roman'
                        work_sheet.Range(cell).Font.Size = 14
                        work_sheet.Range(cell).HorizontalAlignment = win32com.client.constants.xlCenter
                        work_sheet.Range(cell).VerticalAlignment = win32com.client.constants.xlCenter
                        log_file.write(u'{0}\n'.format(value).encode('cp1251'))
            elif re.findall(re_pattern2, value):
                additional_count = re.findall(re_pattern2, value)[0][0]
                doc_number = re.findall(re_pattern2, value)[0][1]
                for path_to_file in os.listdir(path_to_scan):
                    if re.findall(re.compile(u'Доп\. №{0}(.)*{1}'.format(additional_count, doc_number)), path_to_file):
                        work_sheet.Hyperlinks.Add(Anchor=work_sheet.Range(cell),
                                                  Address=os.path.join(path_to_scan, path_to_file))
                        work_sheet.Range(cell).Font.Name = 'Times New Roman'
                        work_sheet.Range(cell).Font.Size = 14
                        work_sheet.Range(cell).HorizontalAlignment = win32com.client.constants.xlCenter
                        work_sheet.Range(cell).VerticalAlignment = win32com.client.constants.xlCenter
                        log_file.write(u'{0}\n'.format(value).encode('cp1251'))
                    elif re.findall(re.compile(u'Д{0} {1}'.format(additional_count, doc_number)), path_to_file):
                        work_sheet.Hyperlinks.Add(Anchor=work_sheet.Range(cell),
                                                  Address=os.path.join(path_to_scan, path_to_file))
                        work_sheet.Range(cell).Font.Name = 'Times New Roman'
                        work_sheet.Range(cell).Font.Size = 14
                        work_sheet.Range(cell).HorizontalAlignment = win32com.client.constants.xlCenter
                        work_sheet.Range(cell).VerticalAlignment = win32com.client.constants.xlCenter
                        log_file.write(u'{0}\n'.format(value).encode('cp1251'))
                    elif re.findall(re.compile(u'{0} Д{1}'.format(doc_number, additional_count)), path_to_file):
                        work_sheet.Hyperlinks.Add(Anchor=work_sheet.Range(cell),
                                                  Address=os.path.join(path_to_scan, path_to_file))
                        work_sheet.Range(cell).Font.Name = 'Times New Roman'
                        work_sheet.Range(cell).Font.Size = 14
                        work_sheet.Range(cell).HorizontalAlignment = win32com.client.constants.xlCenter
                        work_sheet.Range(cell).VerticalAlignment = win32com.client.constants.xlCenter
                        log_file.write(u'{0}\n'.format(value).encode('cp1251'))
            elif re.findall(re_pattern3, value):
                additional_count = re.findall(re_pattern2, value)[0][1]
                doc_number = re.findall(re_pattern2, value)[0][0]
                for path_to_file in os.listdir(path_to_scan):
                    if re.findall(re.compile(u'Доп\. №{0}(.)*{1}'.format(additional_count, doc_number)), path_to_file):
                        work_sheet.Hyperlinks.Add(Anchor=work_sheet.Range(cell),
                                                  Address=os.path.join(path_to_scan, path_to_file))
                        work_sheet.Range(cell).Font.Name = 'Times New Roman'
                        work_sheet.Range(cell).Font.Size = 14
                        work_sheet.Range(cell).HorizontalAlignment = win32com.client.constants.xlCenter
                        work_sheet.Range(cell).VerticalAlignment = win32com.client.constants.xlCenter
                        log_file.write(u'{0}\n'.format(value).encode('cp1251'))
                    elif re.findall(re.compile(u'Д{0} {1}'.format(additional_count, doc_number)), path_to_file):
                        work_sheet.Hyperlinks.Add(Anchor=work_sheet.Range(cell),
                                                  Address=os.path.join(path_to_scan, path_to_file))
                        work_sheet.Range(cell).Font.Name = 'Times New Roman'
                        work_sheet.Range(cell).Font.Size = 14
                        work_sheet.Range(cell).HorizontalAlignment = win32com.client.constants.xlCenter
                        work_sheet.Range(cell).VerticalAlignment = win32com.client.constants.xlCenter
                        log_file.write(u'{0}\n'.format(value).encode('cp1251'))
                    elif re.findall(re.compile(u'{0} Д{1}'.format(doc_number, additional_count)), path_to_file):
                        work_sheet.Hyperlinks.Add(Anchor=work_sheet.Range(cell),
                                                  Address=os.path.join(path_to_scan, path_to_file))
                        work_sheet.Range(cell).Font.Name = 'Times New Roman'
                        work_sheet.Range(cell).Font.Size = 14
                        work_sheet.Range(cell).HorizontalAlignment = win32com.client.constants.xlCenter
                        work_sheet.Range(cell).VerticalAlignment = win32com.client.constants.xlCenter
                        log_file.write(u'{0}\n'.format(value).encode('cp1251'))
            elif re.findall(re_pattern, value):
                doc_number = re.findall(re_pattern, value)[0]
                for path_to_file in os.listdir(path_to_scan):
                    if re.findall(re.compile(u'{0}'.format(doc_number)), path_to_file):
                        work_sheet.Hyperlinks.Add(Anchor=work_sheet.Range(cell),
                                                  Address=os.path.join(path_to_scan, path_to_file))
                        work_sheet.Range(cell).Font.Name = 'Times New Roman'
                        work_sheet.Range(cell).Font.Size = 14
                        work_sheet.Range(cell).HorizontalAlignment = win32com.client.constants.xlCenter
                        work_sheet.Range(cell).VerticalAlignment = win32com.client.constants.xlCenter
                        log_file.write(u'{0}\n'.format(value).encode('cp1251'))
            # print(u'{0}  {1}'.format(value, re.findall(re_pattern, value)))
        # if hyperlink:
        #     if path_to_scan not in hyperlink:
        #         hyperlink = os.path.join(path_to_scan, hyperlink)
        #         work_sheet.Range(cell).Hyperlinks(1).Address = hyperlink
        row += 1
    # print(len(hyperlinks_dict.keys()))
    # print(len([i for i in hyperlinks_dict.values() if i[1]]))
    # good_path = os.path.join(path_to_scan, work_sheet.Range('B2').Hyperlinks(1).Address)
    # work_sheet.Range('B2').Hyperlinks(1).Address = good_path
    work_book.Save()
    work_book.Close()
except pywintypes.com_error as e:
    print(e.args[1].decode('cp1251'))
    print(e.args[2][2])
except Exception as e:
    print(e)
finally:
    excel.Application.Quit()
    log_file.close()


# Sub Макрос1()
# '
# ' Макрос1 Макрос
# '
#
# '
#     ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
#         "Корреспонденция%20ОТД.%20№43.xlsx"
#     Range("A2").Select
#     ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="excel_updater.py"
#     Selection.Hyperlinks(1).Address = ".idea"
# End Sub

# Sub Макрос3()
# '
# ' Макрос3 Макрос
# '
#
# '
#     ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
#         "Лист%20Microsoft%20Excel.xlsx"
# End Sub
