from App import App
from StatementData import StatementData

"""
Заранее определим некоторые переменные
"""
#  Файл с данными по тональности
file_with_data = "data/excel_file_temp.xlsx"

#  Файл с информацией по обращениям 
#  (номер обращения, исполнитель, инициатор, email)
all_application_file = "data/all_application.xlsx"

#  Цвет выделенного мышью фрагмента диаграммы
click_color = "#00a4ef"

#  Цвета фрагментов круговой диаграммы. Порядок цветов важен.
#  1 - Без комментариев
#  2 - Нейтральное отношение
#  3 - Неудовлетворительное
#  4 - Полная удовлетворённость
pie_colors = ["#737373", "#FFB900", "#F25022", "#7FBA00"]

#  Путь к фоновому изображению
background_image = "images\\background_image_radar.png"


if __name__ == "__main__":
    print("Подключил необходимые библиотеки")

    # Создаём набор данных
    my_data = StatementData(file_with_data, all_application_file)

    # Создаём экземпляр самого приложения
    app = App(my_data, background_image)
    #  Задаём размер окна. Сначала ширина, затем высота.
    app.geometry("1280x720")
    app.maxsize(1280, 720)
    app.configure(background='white')
    app.mainloop()
    print("Завершение работы")
