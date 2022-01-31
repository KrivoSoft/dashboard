from openpyxl import load_workbook
import pandas as pd
import time

class StatementData:
    """
    Класс для датафрейма и оперирования над данными
    """

    file_with_data = ""  # Имя файла с данными
    last_modified_time = None  # Дата изменения файла с данными
    dataframe = None  # Весь датафрейм - таблица Excel-файла

    summary_performer = []  # Датафрейм сводной таблицы по тональности заявителя
    quantity_performer = []  # Данные сводной таблицы по тональности исполнителя
    labels_performer = []  # Заголовки сводной таблицы по тональности исполнителя

    summary_requester = []  # Датафрейм сводной таблицы по тональности заявителя
    quantity_requester = []  # Данные сводной по тональности заявителя
    labels_requester = []  # Заголовки сводной таблицы по тональности исполнителя

    all_application_file = ""  # Файл с информацией по всем обращениям
    bad_app_list = []  # Список последних 10 обращений
    #                             с неудовлетворительной тональностью исполнителя

    all_app_data = None  # Датафрейм со всеми обращениями
    info_bad_app = []  # Список с информацией о всех обращениях

    #                             с невежливым обращением.

    def __init__(self, filename, all_app_file):
        self.file_with_data = filename
        self.all_application_file = all_app_file
        self.get_summary_info()

    def get_summary_info(self):
        """
        Получаем сводные таблицы
        """
        try:
            self.dataframe = self.get_data_from_file(self.file_with_data)  # Подгрузим данные
        except FileNotFoundError:
            print("Не могу найти файл: " + self.file_with_data)
            exit(1)

        self.dataframe = self.clear_data(self.dataframe)

        #  Получаем последние 10 обращений с неудовлетворённой
        #  тональностью заявителя
        bad_application = self.dataframe.loc[self.dataframe['Тональность исполнителя'] == 'Невежливое обращение']
        bad_application = bad_application.tail(10)
        bad_app_list = bad_application['ID Обращения'].tolist()
        self.bad_app_list = bad_app_list

        self.summary_performer = pd.pivot_table(self.dataframe,
                                                columns="Тональность исполнителя",
                                                values="Количество",
                                                aggfunc="sum")

        self.summary_requester = pd.pivot_table(self.dataframe,
                                                index="Тональность заявителя",
                                                aggfunc="sum")

        self.all_app_data = self.get_data_from_file(self.all_application_file)
        self.all_app_data = self.clear_data(self.all_app_data)
        self.info_bad_app = self.get_all_bad_app(self.all_app_data, bad_app_list)

    def get_data_from_file(self, name_of_file):
        """
        Функция чтения данных из Excel-файла
        """

        #  Установим счётчик времени
        start_time = time.time()
        wb = load_workbook(filename=name_of_file, read_only=True)
        ws = wb.active

        #  Превращаем данные в Dataframe
        data = pd.DataFrame(data=ws.values)

        #  Закрытие книги
        wb.close()

        print("Загружены даные из файла ", name_of_file)
        print("Время выполнения: ", time.time() - start_time)

        return data

    def clear_data(self, data):
        """
        Функция предварительной очистки данных
        """

        # На случай, если понадобится удалить лишние столбцы
        # data.drop(['Решение', 'Комментарий заявителя'], axis=1, inplace=True)

        data.columns = data.iloc[0]  # Устанавливаем заголовок столбцов таблицы - значения первой строки
        data.drop(0, axis=0, inplace=True)  # А первую строку удаляем, т.к. там по-прежнему заголовки столбцов

        data['Количество'] = 1  # Добавляем справа столбец "Количество",
        #                                        в котором все ячейки содерат значение 1
        #                                        (необходимо для расчёта сводной таблицы)

        return data

    def convert_data_for_diagrams(self, summary_table):
        """
        Функция для преобразования данных сводной таблицы
        в данные для построения графиков
        """

        #  Получаем названия полей
        labels = []
        for label in summary_table.index:
            labels.append(label)

        #  Получаем массив значений второго столбца
        quantity = summary_table['Количество'].tolist()

        return labels, quantity

    def get_all_bad_app(self, data, list_with_app):

        all_app = []

        for one_app in list_with_app:
            try:
                row = data.loc[data['Номер обращения'] == one_app].values.flatten().tolist()
                if row:
                    all_app.append(row)
            except KeyError:
                print("Не могу найти такой номер обращения: " + one_app)

        return all_app