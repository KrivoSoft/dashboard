from tkinter import *
from tkinter.ttk import Treeview
import os.path
import time
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np


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
        fig_1 = plt.Figure(figsize=(6, 4))

        #  Создание сетки для построения диаграммы 1х1 для 1 графика.
        #  axes - область, где отображается диаграмма
        ax_1 = fig_1.add_subplot(111)

        #  Объект FigureCanvasTkAgg объединяет объект figure и
        #  графическое окно Tkinter (объект Canvas в Tkinter)
        canvas_pie = FigureCanvasTkAgg(fig_1, self)

        #  Размещение элемента в графическом окне через метод place()
        #  place() - абсолютное позиционирование
        canvas_pie.get_tk_widget().place(x=30, y=150)

        #  Создаём круговую диаграмму по тональности заявителя в области axes
        #  patches - хранит клиновидные фрагменты диаграммы
        patches, texts, autotexts = ax_1.pie(data.summary_requester["Количество"],
                                             startangle=90,
                                             autopct='%1.1f%%',
                                             colors=self.pie_colors,
                                             wedgeprops=dict(width=0.5),
                                             pctdistance=0.8,
                                             labeldistance=0.8)

        pie_labels = ["Без коммент.", "Нейтральное", "Неудовл.", "Полная удовл."]

        bbox_props = dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72)
        kw = dict(arrowprops=dict(arrowstyle="-"),
                  bbox=bbox_props, zorder=0, va="center")

        for i, p in enumerate(patches):
            ang = (p.theta2 - p.theta1) / 2. + p.theta1
            y = np.sin(np.deg2rad(ang))
            x = np.cos(np.deg2rad(ang))
            horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
            connectionstyle = "angle,angleA=0,angleB={}".format(ang)
            kw["arrowprops"].update({"connectionstyle": connectionstyle})
            ax_1.annotate(pie_labels[i], xy=(x, y), xytext=(1.35 * np.sign(x), 1.4 * y),
                          horizontalalignment=horizontalalignment, **kw)

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
        fig_2 = plt.Figure(figsize=(4, 4))

        #  Создание Axes
        ax_2 = fig_2.add_subplot(111)

        #  Объединяем фигуру и окно программы
        bar_canvas = FigureCanvasTkAgg(fig_2, self)

        #  Задаём позицию элемента
        bar_canvas.get_tk_widget().place(x=750, y=150)

        #  Построение столбчатой диаграммы по тональности исполнителя
        my_bar = data.summary_performer.plot(kind='bar',
                                             ax=ax_2,
                                             subplots=False,
                                             color=["#7FBA00", "#F25022"],
                                             width=0.25,
                                             legend=False)

        for i in range(len(my_bar.containers)):
            my_bar.bar_label(my_bar.containers[i], label_type='edge')

        # ax_2.legend(fancybox=True, framealpha=0.4, shadow=True, borderpad=1)
        box = ax_2.get_position()
        ax_2.set_position([box.x0, box.y0 + box.height * 0.1,
                           box.width, box.height * 0.9])

        # Put a legend below current axis
        ax_2.legend(loc='upper center', bbox_to_anchor=(0.5, -0.05),
                    fancybox=True,
                    shadow=True,
                    ncol=1,
                    frameon=False)
        ax_2.axis('off')

        self.fig_2 = fig_2
        self.ax_2 = ax_2

        fig_2.canvas.mpl_connect('button_press_event', self.onclick)

        #  Заголовки диаграмм
        Label(self, text="Тональность заявителя", bg="white", font=(None, 16)).place(x=250, y=120)
        Label(self, text="Тональность исполнителя", bg="white", font=(None, 16)).place(x=850, y=120)

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
        # a = event.artist
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
