import tkinter as tk
from tkinter import ttk
import pandas as pd
import time
import os.path
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

"""
Заранее определим некоторые переменные
"""
#  Файл с данными по тональности
file_with_data = "excel_file_temp.xlsx"

#  Файл с информацией по обращениям 
#  (номер обращения, исполнитель, инициатор, email)
all_application_file = "all_application.xlsx"

#  Цвет выделенного мышью фрагмента диаграммы
click_color = "#F71735"
    
    
class App(tk.Tk):
    """
    Класс графического приложения
    """
    
    #  Элементы графического окна приложения
    
    #  Круговая диаграмма
    fig_1 = None
    ax_1 = None
    patches_pie = None
    canvas_pie = None
    
    #  Столбчатая диаграмма
    fig_2 = None
    ax_2 = None
    patches_bar = None
    
    # Список обращений с неудовлетворённой тональностью
    sd_listbox = None
    bad_app_frame = None
    
    #  Таблица с данными об обращениях
    tree = None
    
    def __init__(self, data):
        super().__init__()
        
        self.title('Radar')
        ########################
        #  Круговая диаграмма  #
        ########################
        
        #  Создание объекта типа figure
        #  figure - контейнер самого верхнего уровня 
        fig_1 = plt.Figure(figsize=(3, 3))
        
        #  Создание сетки для построения диаграммы 1х1 для 1 графика.
        #  axes - область, где отображается диаграмма
        ax_1 = fig_1.add_subplot(111)  
        
        #  Объект FigureCanvasTkAgg объединяет объект figure и 
        #  графическое окно Tkinter (объект Canvas в Tkinter)
        canvas_pie = FigureCanvasTkAgg(fig_1, self)
        
        #  Размещение элемента в графическом окне через метод place()
        #  place() - абсолютное позиционирование
        canvas_pie.get_tk_widget().place(x=20, y=10)
        
        #  Создаём круговую диаграмму в области axes
        #  patches - хранит клиновидные фрагменты диаграммы
        patches, texts, autotexts = ax_1.pie(data.summary_requester["Количество"], startangle=90, 
            autopct='%1.1f%%')
        ax_1.set_title("Тональность заявителя")
        
        for p in patches:
            p.set_gid(p.get_facecolor())
            # Активируем выделение
            p.set_picker(True)
        
        self.patches_pie = patches
        
        fig_1.canvas.mpl_connect('pick_event', self.onclick)
        
        self.ax_1 = ax_1
        self.fig_1 = fig_1
        
        #########################
        #  Столбчатая диаграмма #
        #########################
        
        #  Cоздание фигуры
        fig_2 = plt.Figure(figsize=(3, 3.3))
        
        #  Создание Axes
        ax_2 = fig_2.add_subplot(111)
        
        #  Объединяем фигуру и окно программы
        bar_canvas = FigureCanvasTkAgg(fig_2, self)
        
        #  Задаём позицию элемента
        bar_canvas.get_tk_widget().place(x=350, y=10)
        
        #  Построение столбчатой диаграммы
        data.summary_performer.plot(kind='bar', ax=ax_2, subplots=False, rot=0, color=['#5cb85c', '#d9534f'], width=0.08)
        ax_2.set_title("Тональность исполнителя")
        #ax_2.legend(fancybox=True, framealpha=0.4, shadow=True, borderpad=1)
        
        self.fig_2 = fig_2
        self.ax_2 = ax_2
        
        #########################################################
        #  Список обращений с неудовлетворительной тональностью #
        #########################################################
        
        #  Элемент ListBox
        bad_app_frame = tk.LabelFrame(self, text='Неудовл. исполнителя')
        bad_app_frame.place(x=800, y=10)
        sd_listbox = tk.Listbox(bad_app_frame, height=10)
        sd_listbox.grid(row=1, column=1)
        
        #  Вставка номеров обращений с неудовл. тональностью
        #  в список в окне приложения
        for application_number in data.bad_app_list:
            sd_listbox.insert(tk.END, application_number)
        
        self.bad_app_frame = bad_app_frame
        self.sd_listbox = sd_listbox
        
        ###########
        # Таблица #
        ###########
        tree = ttk.Treeview(self, column=("Номер обращения", "Исполнитель", "Инициатор", "Email"), show='headings', height=10)
        tree.column("# 1")
        tree.heading("# 1", text="Номер обращения")
        tree.column("# 2")
        tree.heading("# 2", text="Исполнитель")
        tree.column("# 3")
        tree.heading("# 3", text="Инициатор")
        tree.column("# 4")
        tree.heading("# 4", text="Email")

        # Вставка данных в таблицу
        all_bad_app = data.info_bad_app
        
        for row in all_bad_app:
            tree.insert('', 'end', text="1", values=(row[0], row[1], row[2], row[3]))

        tree.place(x=20, y=400)
        self.tree = tree
        
        #  Запоминаем дату изменения файла
        data.last_modified_time = time.ctime(os.path.getmtime(data.file_with_data))
        
        self.refresh(canvas_pie, bar_canvas, tree, data)
        
        
    def refresh(self, inedible_pie, bar, tree, data):
        """
        Метод обновления приложения
        """
        
        try:
            modified_time = time.ctime(os.path.getmtime(data.file_with_data))
            
            if data.last_modified_time != modified_time:
                
                print("Обновляю данные")
                data.last_modified_time = modified_time
                data.get_summary_info()
                
                self.fig_1.clf()
                self.fig_1.canvas.draw()
                ax_2 = self.ax_2
                
                self.ax_1 = self.fig_1.add_subplot(111)
                self.patches_pie, texts, autotexts = self.ax_1.pie(data.summary_requester["Количество"], startangle=90, autopct='%1.1f%%')
                self.ax_1.set_title("Тональность заявителя")
                
                for p in self.patches_pie:
                    p.set_gid(p.get_facecolor())
                    # Активируем выделение
                    p.set_picker(True)
                
                self.canvas_pie = FigureCanvasTkAgg(self.fig_1, self)
                self.canvas_pie.get_tk_widget().place(x=20, y=10)
                
                #fig_1.canvas.mpl_connect('pick_event', self.onclick)
                self.fig_1.canvas.draw()
                
                fig_2 = self.fig_2
                #ax_2 = fig_2.add_subplot()
                bar_canvas = FigureCanvasTkAgg(fig_2, self)
                #bar.get_tk_widget().grid(row=1, column=1, padx=5, pady=5)
                data.summary_performer.plot(kind='bar', ax=ax_2, subplots=False, rot=0, color=['#5cb85c', '#d9534f'], width=0.08)
                #ax_2.legend(fancybox=True, framealpha=0.4, shadow=True, borderpad=1)
                
                #  Очистка списка с обращениями с невежливой тональностью
                sd_listbox = self.sd_listbox
                sd_listbox.delete(0,tk.END)
                
                for application_number in data.bad_app_list:
                    sd_listbox.insert(tk.END, application_number)
                
                self.sd_listbox = sd_listbox
                
                #  Очистка таблицы
                self.tree.delete(*tree.get_children())
                tree = self.tree
                
                all_bad_app = data.info_bad_app
        
                for row in all_bad_app:
                    tree.insert('', 'end', text="1", values=(row[0], row[1], row[2], row[3]))
                    
                self.tree = tree
                
                print("Построил новую диаграмму.")
            else:
                print("Файл не изменён. Перерасчёт не требуется.")
            
        except FileExistsError:
            print("Файл используется")
        
        self.update()
        self.after(5000, self.refresh, inedible_pie, bar, tree, data)
        
        
    def onclick(self, event):
        #  Действия при нажатии мышкой на фрагмент диаграммы
        for p in self.patches_pie:
            p.set_facecolor(p.get_gid())
        a = event.artist
        # print('on pick:', a, a.get_gid())
        a.set_facecolor(click_color)
        print("Сработало событие выделения фрагмента пирога")
        self.fig_1.canvas.draw()


class Statement_Data():
    '''
    Класс для датафрейма и оперирования над данными
    '''
    
    file_with_data = ""        #  Имя файла с данными
    last_modified_time = None  #  Дата изменения файла с данными
    dataframe = None           #  Весь датафрейм - таблица Excel-файла
    
    summary_performer = []     #  Датафрейм сводной таблицы по тональности заявителя
    quantity_performer = []    #  Данные сводной таблицы по тональности исполнителя
    labels_performer = []      #  Заголовки сводной таблицы по тональности исполнителя
    
    summary_requester = []     #  Датафрейм сводной таблицы по тональности заявителя
    quantity_requester = []    #  Данные сводной по тональности заявителя
    labels_requester = []      #  Заголовки сводной таблицы по тональности исполнителя
    
    all_application_file = ""  #  Файл с информацией по всем обращениям
    bad_app_list = []          #  Список последних 10 обращений
    #                             с неудовлетворительной тональностью исполнителя
    
    all_app_data = None        #  Датафрейм со всеми обращениями
    info_bad_app = []          #  Список с информацией о всех обращениях
    #                             с невежливым обращением.
    
    
    def __init__(self, filename):
        self.file_with_data = filename
        self.all_application_file = all_application_file
        self.get_summary_info()
    
    
    def get_summary_info(self):
        ''' 
        Получаем сводные таблицы
        '''
        try:
            self.dataframe = self.get_data_from_file(self.file_with_data) #  Подгрузим данные
        except(FileNotFoundError):
            print("Не могу найти файл: " + file_with_data)
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
        '''
        Функция чтения данных из Excel-файла
        '''
        
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
        ''' 
        Функция предварительной очистки данных
        '''
        
        # На случай, если понадобится удалить лишние столбцы
        #data.drop(['Решение', 'Комментарий заявителя'], axis=1, inplace=True)
        
        data.columns = data.iloc[0] #            Устанавливаем заголовок столбцов таблицы - значения первой строки
        data.drop(0, axis=0, inplace=True) #     А первую строку удаляем, т.к. там по-прежнему заголовки столбцов
        
        data['Количество'] = 1 #                 Добавляем справа столбец "Количество", 
        #                                        в котором все ячейки содерат значение 1 
        #                                        (необходимо для расчёта сводной таблицы)
        
        return data
        

    def convert_data_for_diagrams(self, summary_table):
        ''' 
        Функция для преобразования данных сводной таблицы
        в данные для построения графиков
        '''
        
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
                if row != []:
                    all_app.append(row)
            except KeyError:
                print("Не могу найти такой номер обращения: " + one_app)
        
        return all_app


if __name__ == "__main__":
    print("Подключил необходимые библиотеки")
    
    my_data = Statement_Data(file_with_data)
    
    app = App(my_data)
    
    #  Задаём размер окна. Сначала ширина, затем высота.
    app.geometry("1000x600")
    app.configure(background='white')
    app.mainloop()
    print("Завершение работы")
