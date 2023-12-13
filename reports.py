import openpyxl
from openpyxl.styles import Font
from openpyxl.workbook import Workbook
import db_manage as db
import excel_processing
import locale

locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')


class OriginalReport:
    def __init__(self, file_name):
        # self.filename = file_name
        # self.excel_book = openpyxl.open(file_name, read_only=True)
        self.excel_book = excel_processing.ExcelBook(file_name)
        self.file_name = self.excel_book.filename
        self.sample = None

        # Grand_Pharm
        if self.excel_book.get_cell_value('A2') == 'Наименование':
            self.sample = 'Grand_Pharm'

        # Pharm_Luxe (Intellia)
        elif self.excel_book.get_cell_value('D3') == 'Наименование':
            self.sample = 'Pharm_Luxe'

        # Whole_Pharm (Genex)
        elif self.excel_book.get_cell_value('B5') == 'Наименование':
            self.sample = 'Whole_Pharm'

        # Whole_Pharm_2 (Genex)
        elif self.excel_book.get_cell_value('B4') == 'Товар':
            self.sample = 'Whole_Pharm_2'

        # Meros
        elif self.excel_book.get_cell_value('B3', sheet_index=1) \
                == 'Наименование':
            self.sample = 'Meros'

        # Young_Pharm
        elif self.excel_book.get_cell_value('A1') == 'Артикул':
            self.sample = 'Young_Pharm'

    def get_report_dict(self):
        report_dict = {}
        exception_unique_names = []

        # Grand_Pharm
        if self.sample == 'Grand_Pharm':
            for n in range(3, self.excel_book.get_max_row() + 1):
                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"A{n}"))
                except:
                    name = self.excel_book.get_cell_value(f"A{n}")
                    exception_unique_names.append(name)
                    continue

                # record in reports
                balance_beginning = int(
                    self.excel_book.get_cell_value(f"C{n}")) + \
                                    int(self.excel_book.get_cell_value(
                                        f"D{n}"))
                balance_end = int(self.excel_book.get_cell_value(f"I{n}"))
                consumption = balance_beginning - balance_end
                if name not in report_dict:
                    new_data = {
                        'balance_beginning': balance_beginning,
                        'balance_beginning_price': balance_beginning * price,

                        'consumption': consumption,
                        'consumption_price': consumption * price,

                        'balance_end': balance_end,
                        'balance_end_price': balance_end * price
                    }
                    report_dict[name] = new_data
                else:
                    report_dict[name][
                        'balance_beginning'] += balance_beginning
                    report_dict[name][
                        'balance_beginning_price'] += balance_beginning * price

                    report_dict[name]['consumption'] += consumption
                    report_dict[name][
                        'consumption_price'] += consumption * price

                    report_dict[name]['balance_end'] += balance_end
                    report_dict[name][
                        'balance_end_price'] += balance_end * price

        # Pharm_Luxe (Intellia)
        elif self.sample == 'Pharm_Luxe':
            for n in range(5, self.excel_book.get_max_row() + 1):
                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"D{n}"))
                except:
                    name = self.excel_book.get_cell_value(f"D{n}")
                    exception_unique_names.append(name)
                    continue

                # record in reports
                balance_beginning = int(
                    self.excel_book.get_cell_value(f"K{n}")) + \
                                    int(self.excel_book.get_cell_value(
                                        f"L{n}"))
                balance_end = int(self.excel_book.get_cell_value(f"O{n}"))
                consumption = balance_beginning - balance_end
                if name not in report_dict:
                    new_data = {
                        'balance_beginning': balance_beginning,
                        'balance_beginning_price': balance_beginning * price,

                        'consumption': consumption,
                        'consumption_price': consumption * price,

                        'balance_end': balance_end,
                        'balance_end_price': balance_end * price
                    }
                    report_dict[name] = new_data
                else:
                    report_dict[name]['balance_beginning'] += balance_beginning
                    report_dict[name][
                        'balance_beginning_price'] += balance_beginning * price

                    report_dict[name]['consumption'] += consumption
                    report_dict[name][
                        'consumption_price'] += consumption * price

                    report_dict[name]['balance_end'] += balance_end
                    report_dict[name][
                        'balance_end_price'] += balance_end * price

        # Whole_Pharm (Genex)
        elif self.sample == 'Whole_Pharm':
            for n in range(6, self.excel_book.get_max_row() + 1):
                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"B{n}"))
                except:
                    name = self.excel_book.get_cell_value(f"B{n}")
                    exception_unique_names.append(name)
                    continue

                # record in reports
                try:
                    balance_beginning = int(
                        self.excel_book.get_cell_value(f"G{n}")) + \
                                        int(self.excel_book.get_cell_value(
                                            f"H{n}"))
                    balance_end = int(self.excel_book.get_cell_value(f"K{n}"))
                except ValueError:
                    continue
                consumption = balance_beginning - balance_end
                if name not in report_dict:
                    new_data = {
                        'balance_beginning': balance_beginning,
                        'balance_beginning_price': balance_beginning * price,

                        'consumption': consumption,
                        'consumption_price': consumption * price,

                        'balance_end': balance_end,
                        'balance_end_price': balance_end * price
                    }
                    report_dict[name] = new_data
                else:
                    report_dict[name]['balance_beginning'] += balance_beginning
                    report_dict[name][
                        'balance_beginning_price'] += balance_beginning * price

                    report_dict[name]['consumption'] += consumption
                    report_dict[name][
                        'consumption_price'] += consumption * price

                    report_dict[name]['balance_end'] += balance_end
                    report_dict[name][
                        'balance_end_price'] += balance_end * price

        # Whole_Pharm_2 (Genex)
        elif self.sample == 'Whole_Pharm_2':
            for n in range(5, self.excel_book.get_max_row()):
                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"B{n}"))
                except:
                    name = self.excel_book.get_cell_value(f"B{n}")
                    exception_unique_names.append(name)
                    continue

                # record in reports
                try:
                    balance_beginning = int(
                        self.excel_book.get_cell_value(f"L{n}")
                    ) + int(
                        self.excel_book.get_cell_value(f"M{n}")
                    )

                    balance_end = int(
                        self.excel_book.get_cell_value(f"Q{n}")
                    )

                except ValueError:
                    continue

                consumption = balance_beginning - balance_end

                if name not in report_dict:
                    new_data = {
                        'balance_beginning': balance_beginning,
                        'balance_beginning_price': balance_beginning * price,

                        'consumption': consumption,
                        'consumption_price': consumption * price,

                        'balance_end': balance_end,
                        'balance_end_price': balance_end * price
                    }
                    report_dict[name] = new_data

                else:
                    report_dict[name]['balance_beginning'] += balance_beginning
                    report_dict[name]['balance_beginning_price'] += \
                        balance_beginning * price

                    report_dict[name]['consumption'] += consumption
                    report_dict[name]['consumption_price'] += \
                        consumption * price

                    report_dict[name]['balance_end'] += balance_end
                    report_dict[name]['balance_end_price'] += \
                        balance_end * price

        # Meros
        elif self.sample == 'Meros':
            for n in range(5, self.excel_book.get_max_row(sheet_index=1) + 1):
                # get name
                name_value = self.excel_book.get_cell_value(f"B{n}",
                                                            sheet_index=1)
                if name_value != 0:
                    try:
                        name, price = db.get_general_name_and_price(name_value)
                    except:
                        name = name_value
                        exception_unique_names.append(name)
                        continue
                    # record in reports
                    try:
                        balance_beginning = int(
                            self.excel_book.get_cell_value(f"D{n}",
                                                           sheet_index=1)) + \
                                            int(self.excel_book.get_cell_value(
                                                f"E{n}", sheet_index=1))
                        balance_end = int(
                            self.excel_book.get_cell_value(f"G{n}",
                                                           sheet_index=1))
                    except ValueError:
                        continue
                    consumption = balance_beginning - balance_end
                    if name not in report_dict:
                        new_data = {
                            'balance_beginning': balance_beginning,
                            'balance_beginning_price': balance_beginning * price,

                            'consumption': consumption,
                            'consumption_price': consumption * price,

                            'balance_end': balance_end,
                            'balance_end_price': balance_end * price
                        }
                        report_dict[name] = new_data
                    else:
                        report_dict[name][
                            'balance_beginning'] += balance_beginning
                        report_dict[name][
                            'balance_beginning_price'] += balance_beginning * price

                        report_dict[name]['consumption'] += consumption
                        report_dict[name][
                            'consumption_price'] += consumption * price

                        report_dict[name]['balance_end'] += balance_end
                        report_dict[name][
                            'balance_end_price'] += balance_end * price

        # Navbahor
        elif self.sample == 'Navbahor':
            for n in range(2, self.excel_book.get_max_row() + 1):
                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"B{n}"))
                except:
                    name = self.excel_book.get_cell_value(f"B{n}")
                    exception_unique_names.append(name)
                    continue

                # record in reports
                try:
                    balance_end = self.excel_book.get_cell_value(f"M{n}")
                    consumption = self.excel_book.get_cell_value(f"L{n}")
                except ValueError:
                    continue
                balance_beginning = consumption + balance_end
                if name not in report_dict:
                    new_data = {
                        'balance_beginning': balance_beginning,
                        'balance_beginning_price': balance_beginning * price,

                        'consumption': consumption,
                        'consumption_price': consumption * price,

                        'balance_end': balance_end,
                        'balance_end_price': balance_end * price
                    }
                    report_dict[name] = new_data
                else:
                    report_dict[name]['balance_beginning'] += balance_beginning
                    report_dict[name][
                        'balance_beginning_price'] += balance_beginning * price

                    report_dict[name]['consumption'] += consumption
                    report_dict[name][
                        'consumption_price'] += consumption * price

                    report_dict[name]['balance_end'] += balance_end
                    report_dict[name][
                        'balance_end_price'] += balance_end * price

        # Oxy_Med
        elif self.sample == 'Oxy_Med':
            for n in range(5, self.excel_book.get_max_row() + 1):
                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"D{n}"))
                except:
                    name = self.excel_book.get_cell_value(f"D{n}")
                    exception_unique_names.append(name)
                    continue

                # record in reports
                try:
                    balance_beginning = self.excel_book.get_cell_value(f"F{n}")
                    consumption = self.excel_book.get_cell_value(f"G{n}")
                    balance_end = self.excel_book.get_cell_value(f"H{n}")
                except ValueError:
                    continue
                if name not in report_dict:
                    new_data = {
                        'balance_beginning': balance_beginning,
                        'balance_beginning_price': balance_beginning * price,

                        'consumption': consumption,
                        'consumption_price': consumption * price,

                        'balance_end': balance_end,
                        'balance_end_price': balance_end * price
                    }
                    report_dict[name] = new_data
                else:
                    report_dict[name]['balance_beginning'] += balance_beginning
                    report_dict[name][
                        'balance_beginning_price'] += balance_beginning * price

                    report_dict[name]['consumption'] += consumption
                    report_dict[name][
                        'consumption_price'] += consumption * price

                    report_dict[name]['balance_end'] += balance_end
                    report_dict[name][
                        'balance_end_price'] += balance_end * price

        # Pharma_cosmos
        elif self.sample == 'Pharma_cosmos':
            for n in range(7, self.excel_book.get_max_row() + 1):
                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"D{n}"))
                except:
                    name = self.excel_book.get_cell_value(f"D{n}")
                    exception_unique_names.append(name)
                    continue

                # record in reports
                try:
                    balance_beginning = self.excel_book.get_cell_value(
                        f"H{n}") + \
                                        self.excel_book.get_cell_value(f"J{n}")
                    consumption = self.excel_book.get_cell_value(f"N{n}")
                    balance_end = self.excel_book.get_cell_value(f"P{n}")
                except ValueError:
                    continue
                if name not in report_dict:
                    new_data = {
                        'balance_beginning': balance_beginning,
                        'balance_beginning_price': balance_beginning * price,

                        'consumption': consumption,
                        'consumption_price': consumption * price,

                        'balance_end': balance_end,
                        'balance_end_price': balance_end * price
                    }
                    report_dict[name] = new_data
                else:
                    report_dict[name]['balance_beginning'] += balance_beginning
                    report_dict[name][
                        'balance_beginning_price'] += balance_beginning * price

                    report_dict[name]['consumption'] += consumption
                    report_dict[name][
                        'consumption_price'] += consumption * price

                    report_dict[name]['balance_end'] += balance_end
                    report_dict[name][
                        'balance_end_price'] += balance_end * price

        # Asklepiy
        elif self.sample == 'Asklepiy':
            for n in range(2, self.excel_book.get_max_row() + 1):
                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"D{n}"))
                except:
                    name = self.excel_book.get_cell_value(f"D{n}")
                    exception_unique_names.append(name)
                    continue

                # record in reports
                try:
                    consumption = self.excel_book.get_cell_value(f"G{n}")
                except ValueError:
                    continue
                if name not in report_dict:
                    new_data = {
                        'balance_beginning': 0,
                        'balance_beginning_price': 0,

                        'consumption': consumption,
                        'consumption_price': consumption * price,

                        'balance_end': 0,
                        'balance_end_price': 0
                    }
                    report_dict[name] = new_data
                else:
                    report_dict[name]['consumption'] += consumption
                    report_dict[name][
                        'consumption_price'] += consumption * price

        # Pharma_Choice
        elif self.sample == 'Pharma_Choice':
            new_data = {}
            # get data from first sheet
            for n in range(2, self.excel_book.get_max_row() + 1):
                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"B{n}"))
                except:
                    name = self.excel_book.get_cell_value(f"B{n}")
                    while name[-1] == ' ':
                        name = name[:-1]
                    exception_unique_names.append(name)
                    continue
                if name not in new_data:
                    new_data[name] = {
                        'balance_beginning': self.excel_book.get_cell_value(
                            f"H{n}"),
                        'balance_beginning_price': self.excel_book.get_cell_value(
                            f"H{n}") * price,

                        'consumption': 0,
                        'consumption_price': 0,

                        'balance_end': 0,
                        'balance_end_price': 0,
                    }
                else:
                    new_data[name][
                        'balance_beginning'] += self.excel_book.get_cell_value(
                        f"H{n}")
                    new_data[name]['balance_beginning_price'] += \
                        self.excel_book.get_cell_value(f"H{n}") * price

            # get data from second sheet
            for n in range(2, self.excel_book.get_max_row(sheet_index=1) + 1):
                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"D{n}", sheet_index=1))
                except:
                    name = self.excel_book.get_cell_value(f"D{n}",
                                                          sheet_index=1)
                    while name[-1] == ' ':
                        name = name[:-1]
                    exception_unique_names.append(name)
                    continue
                if name not in new_data:
                    new_data[name] = {
                        'balance_beginning': 0,
                        'balance_beginning_price': 0,

                        'consumption':
                            self.excel_book.get_cell_value(f"F{n}",
                                                           sheet_index=1),
                        'consumption_price':
                            self.excel_book.get_cell_value(f"F{n}",
                                                           sheet_index=1) * price,

                        'balance_end': 0,
                        'balance_end_price': 0,
                    }
                else:
                    new_data[name]['consumption'] += \
                        self.excel_book.get_cell_value(f"F{n}", sheet_index=1)
                    new_data[name]['consumption_price'] += \
                        self.excel_book.get_cell_value(f"F{n}",
                                                       sheet_index=1) * price

            for name in new_data:
                new_data[name]['balance_end'] = new_data[name][
                                                    'balance_beginning'] - \
                                                new_data[name]['consumption']
                new_data[name]['balance_end_price'] = \
                    new_data[name]['balance_beginning_price'] - \
                    new_data[name]['consumption_price']
            for name in new_data:
                if name not in report_dict:
                    report_dict[name] = new_data[name]
                else:
                    report_dict[name]['balance_beginning'] += \
                        new_data[name]['balance_beginning']
                    report_dict[name]['balance_beginning_price'] += \
                        new_data[name]['balance_beginning_price']

                    report_dict[name]['consumption'] += \
                        new_data[name]['consumption']
                    report_dict[name]['consumption_price'] += \
                        new_data[name]['consumption_price']

                    report_dict[name]['balance_end'] += new_data[name][
                        'balance_end']
                    report_dict[name]['balance_end_price'] += new_data[name][
                        'balance_end_price']

        # Young_Pharm
        elif self.sample == 'Young_Pharm':
            self.sample = 'Young_Pharm'
            for n in range(3, self.excel_book.get_max_row() + 1):

                # ignore some rows
                if self.excel_book.get_cell_value(f"E{n}") in \
                        ['Гранд Фарм', 'Фарм Люкс Инвест']:
                    continue

                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"A{n}")
                    )
                except:
                    name = self.excel_book.get_cell_value(f"A{n}")
                    exception_unique_names.append(name)
                    continue

                # record in reports
                try:
                    balance_beginning = (
                            self.excel_book.get_cell_value(f"F{n}") +
                            self.excel_book.get_cell_value(f"G{n}")
                    )
                    consumption = self.excel_book.get_cell_value(f"H{n}")
                    balance_end = self.excel_book.get_cell_value(f"I{n}")
                except ValueError:
                    continue
                if name not in report_dict:
                    new_data = {
                        'balance_beginning':
                            balance_beginning,
                        'balance_beginning_price':
                            balance_beginning * price,

                        'consumption':
                            consumption,
                        'consumption_price':
                            consumption * price,

                        'balance_end':
                            balance_end,
                        'balance_end_price':
                            balance_end * price
                    }
                    report_dict[name] = new_data
                else:
                    report_dict[name]['balance_beginning'] \
                        += balance_beginning
                    report_dict[name]['balance_beginning_price'] \
                        += balance_beginning * price

                    report_dict[name]['consumption'] \
                        += consumption
                    report_dict[name]['consumption_price'] \
                        += consumption * price

                    report_dict[name]['balance_end'] \
                        += balance_end
                    report_dict[name]['balance_end_price'] \
                        += balance_end * price

        # Akmal_pharm
        elif self.sample == 'Akmal_pharm':
            for n in range(12, self.excel_book.get_max_row() + 1):
                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"B{n}"))
                except:
                    name = self.excel_book.get_cell_value(f"B{n}")
                    exception_unique_names.append(name)
                    continue
                if name == 0:
                    continue

                # record in reports
                try:
                    balance_beginning = self.excel_book.get_cell_value(
                        f"K{n}") + \
                                        self.excel_book.get_cell_value(f"M{n}")
                    consumption = self.excel_book.get_cell_value(f"O{n}")
                    balance_end = self.excel_book.get_cell_value(f"S{n}")
                except ValueError:
                    continue

                if name not in report_dict:
                    new_data = {
                        'balance_beginning': balance_beginning,
                        'balance_beginning_price': balance_beginning * price,

                        'consumption': consumption,
                        'consumption_price': consumption * price,

                        'balance_end': balance_end,
                        'balance_end_price': balance_end * price
                    }
                    report_dict[name] = new_data
                else:
                    report_dict[name]['balance_beginning'] += balance_beginning
                    report_dict[name][
                        'balance_beginning_price'] += balance_beginning * price

                    report_dict[name]['consumption'] += consumption
                    report_dict[name][
                        'consumption_price'] += consumption * price

                    report_dict[name]['balance_end'] += balance_end
                    report_dict[name][
                        'balance_end_price'] += balance_end * price

        # Best_Pharm
        elif self.sample == 'Best_Pharm':
            new_data = {}
            for n in range(2, self.excel_book.get_max_row() + 1):
                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"D{n}"))
                except:
                    name = self.excel_book.get_cell_value(f"D{n}")
                    exception_unique_names.append(name)
                    continue
                if name == 0:
                    continue

                try:
                    balance_beginning = self.excel_book.get_cell_value(f'F{n}',
                                                                       round_value=False) + \
                                        self.excel_book.get_cell_value(f'G{n}',
                                                                       round_value=False)
                    consumption = - float(
                        self.excel_book.get_cell_value(f'H{n}',
                                                       round_value=False))
                    balance_end = self.excel_book.get_cell_value(f'I{n}',
                                                                 round_value=False)

                    if name not in new_data:
                        new_data[name] = {
                            'balance_beginning': balance_beginning,
                            'balance_beginning_price': balance_beginning * price,

                            'consumption': consumption,
                            'consumption_price': consumption * price,

                            'balance_end': balance_end,
                            'balance_end_price': balance_end * price
                        }
                    else:
                        new_data[name][
                            'balance_beginning'] += balance_beginning
                        new_data[name][
                            'balance_beginning'] += balance_beginning * price

                        new_data[name]['consumption'] += consumption
                        new_data[name]['consumption'] += consumption * price

                        new_data[name]['balance_end'] += balance_end
                        new_data[name]['balance_end'] += balance_end * price
                except ValueError:
                    continue

            for name in new_data:
                if name not in report_dict:
                    report_dict[name] = {}
                    for kw in new_data[name]:
                        report_dict[name][kw] = round(new_data[name][kw])
                else:
                    for kw in new_data[name]:
                        report_dict[name][kw] += round(new_data[name][kw])

        # Zenta_Pharm
        elif self.sample == 'Zenta_Pharm':
            for n in range(6, self.excel_book.get_max_row() + 1):
                # get name
                try:
                    name, price = db.get_general_name_and_price(
                        self.excel_book.get_cell_value(f"B{n}"))
                except Exception as err:
                    name = self.excel_book.get_cell_value(f"B{n}")
                    exception_unique_names.append(name)
                    continue

                balance_beginning = self.excel_book.get_cell_value(f'I{n}') + \
                                    self.excel_book.get_cell_value(f'J{n}')
                consumption = self.excel_book.get_cell_value(f'K{n}')
                balance_end = self.excel_book.get_cell_value(f'L{n}')

                if name not in report_dict:
                    report_dict[name] = {
                        'balance_beginning': balance_beginning,
                        'balance_beginning_price': balance_beginning * price,

                        'consumption': consumption,
                        'consumption_price': consumption * price,

                        'balance_end': balance_end,
                        'balance_end_price': balance_end * price
                    }
                else:
                    report_dict[name]['balance_beginning'] += balance_beginning
                    report_dict[name][
                        'balance_beginning'] += balance_beginning * price

                    report_dict[name]['consumption'] += consumption
                    report_dict[name]['consumption'] += consumption * price

                    report_dict[name]['balance_end'] += balance_end
                    report_dict[name]['balance_end'] += balance_end * price

        if 0 in report_dict:
            del report_dict[0]

        report_dict_sorted_by_group = {
            'C&P': {},
            'OTC': {},
            'RX': {},
            'Eco': {},
            'Gastro': {}
        }

        for name in report_dict:
            group = db.get_group(name)
            report_dict_sorted_by_group[group][name] = report_dict[name]

        for group in ['C&P', 'OTC', 'RX', 'Eco', 'Gastro']:
            if report_dict_sorted_by_group[group] == {}:
                report_dict_sorted_by_group.pop(group)

        if self.sample is not None:
            return report_dict_sorted_by_group, exception_unique_names
        else:
            return None


# print(OriginalReport('Отчет ноябрь.xlsx').get_report_dict()[0])

def make_report_file(*original_report_files: str, new_file_name):
    all_reports_dict = {
        'Total': {}
    }
    sum_numbers = {
        'Total': {}
    }
    all_exception_names = []
    none_sample_files = []
    for report_file in original_report_files:
        # try:
        report = OriginalReport(report_file)
        if report.sample is None:
            none_sample_files.append(report_file)
        else:
            new_report_dict, exception_names = \
                report.get_report_dict()
            all_exception_names += \
                list(set(exception_names))
            print(new_report_dict)

            for group in new_report_dict:
                if group not in all_reports_dict['Total']:
                    all_reports_dict['Total'][group] = {}

                for name in new_report_dict[group]:
                    if name not in all_reports_dict['Total'][group]:
                        all_reports_dict['Total'][group][name] = {}
                        for kw in new_report_dict[group][name]:
                            all_reports_dict['Total'][group][name][kw] = \
                                new_report_dict[group][name][kw]
                            sum_numbers['Total'][kw] = \
                                new_report_dict[group][name][kw]
                    else:
                        for kw in new_report_dict[group][name]:
                            all_reports_dict['Total'][group][name][kw] += \
                                new_report_dict[group][name][kw]
                            sum_numbers['Total'][kw] += \
                                new_report_dict[group][name][kw]

            if report.sample not in all_reports_dict:
                all_reports_dict[report.sample] = {}
                sum_numbers[report.sample] = {}
                for group in new_report_dict:
                    all_reports_dict[report.sample][group] = {}
                    for name in new_report_dict[group]:
                        all_reports_dict[report.sample][group][name] = {}
                        for kw in new_report_dict[group][name]:
                            all_reports_dict[report.sample] \
                                [group][name][kw] = \
                                new_report_dict[group][name][kw]
                            sum_numbers[report.sample][kw] = \
                                new_report_dict[group][name][kw]
            else:
                for group in new_report_dict:
                    for name in new_report_dict[group]:
                        for kw in new_report_dict[group][name]:
                            all_reports_dict[report.sample][group][name] \
                                [kw] += new_report_dict[group][name][kw]
                            sum_numbers[report.sample][kw] += \
                                new_report_dict[group][name][kw]

        # except Exception as err:
        #     print(report_file, 'Exception: ', err)
        #     none_sample_files.append(report_file)

    print(all_reports_dict)

    for sheet_name in all_reports_dict:
        for group in all_reports_dict[sheet_name]:
            group_dict = all_reports_dict[sheet_name][group]
            group_dict_sorted = \
                dict(sorted(group_dict.items(), key=lambda item: item[0]))
            all_reports_dict[sheet_name][group] = group_dict_sorted

    excel_processing.make_report_excel(
        all_reports_dict, sum_numbers, new_file_name
    )
    return all_exception_names, none_sample_files

# report1 = OriginalReport('Sorrento.xlsx')
# report2 = OriginalReport('Отчет Сорренто сент.xlsx')
# report3 = OriginalReport('Alfa Wassermann.xlsx')
#

#
# print(make_report_file('Интеллия.xls',
#                        'Генекс фарма остатка.xlsx',
#                        new_file_name='MYREPORT.xlsx'))


# report_dict = OriginalReport('Report1 11102023094352.xls').get_report_dict()
# print(report_dict)
