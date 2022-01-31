from tkinter import *
from tkinter.ttk import Treeview
import os.path
import time
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class App(Tk):
    """
    Класс графического приложения
    """

    #  Элементы графического окна приложения

    #  Круговая диаграмма
    fig_1 = None
    ax_1 = None
    patches_pie = None
    canvas_pie = None
    pie_colors = ["#737373", "#FFB900", "#F25022", "#7FBA00"]

    #  Столбчатая диаграмма
    fig_2 = None
    ax_2 = None
    patches_bar = None

    # Список обращений с неудовлетворённой тональностью
    sd_listbox = None
    bad_app_frame = None

    #  Таблица с данными об обращениях
    tree = None

    # Фоновое изображение
    bg_label = None
    bg_image = None

    # Общий набор данных
    data = None

    def __init__(self, data, bg_image):
        super().__init__()
        self.bg_image = bg_image
        self.data = data
        self.update_window(data)

    def update_window(self, data):
        """
        Метод обновления основного окна приложения
        """

        #  Данный участок кода будет работать при повторном обновлении данных
        if self.fig_1 is not None:
            #  Очистка всех графических элементов от старых значений
            self.fig_1.clf()
            self.fig_2.clf()
            sd_listbox = self.sd_listbox
            sd_listbox.delete(0, END)
            # self.tree.delete(*tree.get_children())

        #  Настройка фонового изображения главного окна
        bg_image = self.bg_image

        try:
            bg_image = PhotoImage(file=self.bg_image)

            bg_label = self.bg_label
            bg_label = Label(self, image=bg_image)
            bg_label.place(x=0, y=0)
            self.bg_label = bg_label
        except TclError:
            print("Не могу найти изображение для фона")

        self.bg_image = bg_image

        self.title('Radar')

        ########################
        #  Круговая диаграмма  #
        ########################

        #  Создание объекта типа figure
        #  figure - контейнер самого верхнего уровня
        fig_1 = plt.Figure(figsize=(5, 5))

        #  Создание сетки для построения диаграммы 1х1 для 1 графика.
        #  axes - область, где отображается диаграмма
        ax_1 = fig_1.add_subplot(111)

        #  Объект FigureCanvasTkAgg объединяет объект figure и
        #  графическое окно Tkinter (объект Canvas в Tkinter)
        canvas_pie = FigureCanvasTkAgg(fig_1, self)

        #  Размещение элемента в графическом окне через метод place()
        #  place() - абсолютное позиционирование
        canvas_pie.get_tk_widget().place(x=30, y=50)

        #  Создаём круговую диаграмму в области axes
        #  patches - хранит клиновидные фрагменты диаграммы
        patches, texts, autotexts = ax_1.pie(data.summary_requester["Количество"], startangle=90,
                                             autopct='%1.1f%%', colors=self.pie_colors)
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
        fig_2 = plt.Figure(figsize=(5, 5.5))

        #  Создание Axes
        ax_2 = fig_2.add_subplot(111)

        #  Объединяем фигуру и окно программы
        bar_canvas = FigureCanvasTkAgg(fig_2, self)

        #  Задаём позицию элемента
        bar_canvas.get_tk_widget().place(x=460, y=50)

        #  Построение столбчатой диаграммы
        data.summary_performer.plot(kind='bar', ax=ax_2, subplots=False, rot=0, color=['#5cb85c', '#d9534f'],
                                    width=0.08)
        ax_2.set_title("Тональность исполнителя")
        # ax_2.legend(fancybox=True, framealpha=0.4, shadow=True, borderpad=1)

        self.fig_2 = fig_2
        self.ax_2 = ax_2

        """
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
        """

        #  Запоминаем дату изменения файла
        data.last_modified_time = time.ctime(os.path.getmtime(data.file_with_data))

        self.refresh(data)

    def refresh(self, data):

        """
        Метод обновления приложения
        """

        try:
            #  Запоминаем время редактирования файла с данными
            modified_time = time.ctime(os.path.getmtime(data.file_with_data))

            #  Если она изменилась, то обновлением данные
            if data.last_modified_time != modified_time:

                print("Обновляю данные.")
                data.get_summary_info()
                data.last_modified_time = modified_time
                self.update_window(data)
                print("Построил новые диаграммы.")
            else:
                print("Файл не изменён. Перерасчёт не требуется.")

        except FileExistsError:
            print("Файл используется.")

        self.update()
        self.after(5000, self.refresh, data)

    def onclick(self, event):
        """
        Действия при нажатии мышкой на фрагмент диаграммы
        """
        for p in self.patches_pie:
            p.set_facecolor(p.get_gid())
        a = event.artist
        # previous_color = a.get_facecolor()
        # a.set_facecolor(click_color)
        # self.fig_1.canvas.draw()
        # print('on pick:', a, a.get_gid())

        print("Сработало событие выделения фрагмента пирога")
        self.show_table()

    def show_table(self):
        """
        Метод для вывода окна с табличной информацией
        """

        def close_window():
            """
            Метод закрытия окна с таблицей
            """
            window.destroy()

        window = Toplevel(self)
        frame = Frame(window)

        tree = Treeview(window, column=("Номер обращения", "Исполнитель", "Инициатор", "Email"), show='headings',
                        height=10)
        tree.delete(*tree.get_children())
        tree.column("# 1")
        tree.heading("# 1", text="Номер обращения")
        tree.column("# 2")
        tree.heading("# 2", text="Исполнитель")
        tree.column("# 3")
        tree.heading("# 3", text="Инициатор")
        tree.column("# 4")
        tree.heading("# 4", text="Email")

        # Вставка данных в таблицу
        all_bad_app = self.data.info_bad_app

        for row in all_bad_app:
            tree.insert('', 'end', text="1", values=(row[0], row[1], row[2], row[3]))

        #  Добавление кнопки выхода
        button = Button(frame, text="Закрыть", command=close_window)

        tree.pack()
        frame.pack()
        button.pack()
