import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QLineEdit, QPushButton, QComboBox, QTableWidget,
                             QTableWidgetItem, QDateEdit, QTabWidget, QMessageBox, QFileDialog)
from PyQt5.QtCore import QDate
import sqlite3
from fpdf import FPDF


class CorrespondenceApp(QWidget):
    def __init__(self, db_path):
        super().__init__()
        self.db_path = db_path
        self.init_ui()
        self.create_tables()
        self.update_duty_dus_list()

    def init_ui(self):
        self.setWindowTitle("Система учета корреспонденции")

        # Главный макет
        main_layout = QVBoxLayout()

        # Вкладки
        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget)

        # Вкладка для добавления корреспонденции
        self.create_add_correspondence_tab()
        tab_widget.addTab(self.add_correspondence_tab, "Добавлении нагрузки")

        # Вкладка для управления дежурными офицерами
        self.create_manage_duty_dus_tab()
        tab_widget.addTab(self.manage_duty_dus_tab, "Управление ДУС")

        # Вкладка для поиска корреспонденции
        self.create_search_correspondence_tab()
        tab_widget.addTab(self.search_correspondence_tab, "Статистика нагрузки")

        self.setLayout(main_layout)

    def create_add_correspondence_tab(self):
        # Вкладка для добавления корреспонденции
        self.add_correspondence_tab = QWidget()
        layout = QVBoxLayout()

        # Дата
        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel("Дата:"))
        self.date_input = QDateEdit()
        self.date_input.setDate(QDate.currentDate())
        self.date_input.setCalendarPopup(True)
        date_layout.addWidget(self.date_input)
        layout.addLayout(date_layout)

        # Тип корреспонденции
        corr_type_layout = QHBoxLayout()
        corr_type_layout.addWidget(QLabel("Тип информации:"))
        self.corr_type_input = QComboBox()
        self.corr_type_input.addItems([
            "Телефонные переговоры",
            "Телеграфная корреспонденция",
            "Электронные сообщения"
        ])
        corr_type_layout.addWidget(self.corr_type_input)
        layout.addLayout(corr_type_layout)

        # Срочность
        urgency_layout = QHBoxLayout()
        urgency_layout.addWidget(QLabel("Срочность:"))
        self.urgency_input = QComboBox()
        self.urgency_input.addItems([
            "Обыкновенная", "ДСП", "Секретная", "Срочная",
            "Самолет", "Ракета", "Воздух", "Монолит",
            "Секретная срочная", "Секретная самолет", "Секретная монолит",
            "Совершенно секретная", "Особой важности"
        ])
        urgency_layout.addWidget(self.urgency_input)
        layout.addLayout(urgency_layout)

        # Входящая и исходящая корреспонденция
        io_layout = QHBoxLayout()
        io_layout.addWidget(QLabel("Кол-во входящих:"))
        self.incoming_input = QLineEdit()
        io_layout.addWidget(self.incoming_input)
        io_layout.addWidget(QLabel("Кол-во исходящих:"))
        self.outgoing_input = QLineEdit()
        io_layout.addWidget(self.outgoing_input)
        layout.addLayout(io_layout)

        # Период времени
        period_layout = QHBoxLayout()
        period_layout.addWidget(QLabel("Период дежурства:"))
        self.period_input = QComboBox()
        self.period_input.addItems([
            "С 10:00 по 22:00",
            "С 22:00 по 10:00"
        ])
        period_layout.addWidget(self.period_input)
        layout.addLayout(period_layout)

        # Дежурный офицер
        duty_dus_layout = QHBoxLayout()
        duty_dus_layout.addWidget(QLabel("ДУС:"))
        self.duty_dus_input = QComboBox()
        duty_dus_layout.addWidget(self.duty_dus_input)
        layout.addLayout(duty_dus_layout)

        # Кнопка для добавления корреспонденции
        add_button = QPushButton("Добавить")
        add_button.clicked.connect(self.add_correspondence)
        layout.addWidget(add_button)

        self.add_correspondence_tab.setLayout(layout)

    def create_manage_duty_dus_tab(self):
        # Вкладка для управления ДУС
        self.manage_duty_dus_tab = QWidget()
        layout = QVBoxLayout()

        # Добавление дежурного офицера
        add_layout = QVBoxLayout()
        add_layout.addWidget(QLabel("Добавить дежурного по узлу связи"))

        # Создание раскладки для звания
        rank_layout = QHBoxLayout()
        rank_layout.addWidget(QLabel("Звание:"))

        # Создание выпадающего списка для звания
        self.rank_input = QComboBox()
        # Добавление элементов в выпадающий список
        self.rank_input.addItems(["ефр.", "пр-к", "ст. л-т"])

        # Добавление выпадающего списка в раскладку
        rank_layout.addWidget(self.rank_input)

        # Добавление раскладки для звания в основную раскладку
        add_layout.addLayout(rank_layout)

        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("Имя:"))
        self.first_name_input = QLineEdit()
        name_layout.addWidget(self.first_name_input)
        add_layout.addLayout(name_layout)

        last_name_layout = QHBoxLayout()
        last_name_layout.addWidget(QLabel("Фамилия:"))
        self.last_name_input = QLineEdit()
        last_name_layout.addWidget(self.last_name_input)
        add_layout.addLayout(last_name_layout)

        last_last_name_layout = QHBoxLayout()
        last_last_name_layout.addWidget(QLabel("Отчество:"))
        self.last_last_name_input = QLineEdit()
        last_last_name_layout.addWidget(self.last_last_name_input)
        add_layout.addLayout(last_last_name_layout)

        add_button = QPushButton("Добавить")
        add_button.clicked.connect(self.add_duty_dus)
        add_layout.addWidget(add_button)

        layout.addLayout(add_layout)

        # Удаление дежурного офицера
        delete_layout = QVBoxLayout()
        delete_layout.addWidget(QLabel("Удаление дежурного по узлу связи"))

        self.delete_duty_dus_input = QComboBox()
        delete_layout.addWidget(self.delete_duty_dus_input)

        delete_button = QPushButton("Удалить")
        delete_button.clicked.connect(self.delete_duty_dus)
        delete_layout.addWidget(delete_button)

        layout.addLayout(delete_layout)

        self.manage_duty_dus_tab.setLayout(layout)

    def create_search_correspondence_tab(self):
        # Вкладка для поиска корреспонденции
        self.search_correspondence_tab = QWidget()
        layout = QVBoxLayout()

        # Поиск корреспонденции
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Начальная дата:"))
        self.start_date_input = QDateEdit()
        self.start_date_input.setDate(QDate.currentDate())
        self.start_date_input.setCalendarPopup(True)
        search_layout.addWidget(self.start_date_input)

        search_layout.addWidget(QLabel("Конечная дата:"))
        self.end_date_input = QDateEdit()
        self.end_date_input.setDate(QDate.currentDate())
        self.end_date_input.setCalendarPopup(True)
        search_layout.addWidget(self.end_date_input)

        search_button = QPushButton("Поиск")
        search_button.clicked.connect(self.search_correspondence)
        search_layout.addWidget(search_button)

        layout.addLayout(search_layout)

        # Таблица результатов поиска
        self.search_results_table = QTableWidget()
        self.search_results_table.setColumnCount(7)
        self.search_results_table.setHorizontalHeaderLabels([
            "Дата", "Тип", "Срочность", "Входящая", "Исходящая", "Период", "ДУС"
        ])
        layout.addWidget(self.search_results_table)

        # Кнопка для экспорта результатов в PDF
        export_button = QPushButton("Экспортировать в PDF")
        export_button.clicked.connect(self.export_to_pdf)
        layout.addWidget(export_button)

        self.search_correspondence_tab.setLayout(layout)

    def create_tables(self):
        # Создаем таблицы в базе данных, если их еще нет
        try:
            connection = sqlite3.connect(self.db_path)
            cursor = connection.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS duty_dus (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    rank TEXT NOT NULL,
                    first_name TEXT NOT NULL,
                    last_name TEXT NOT NULL,
                    last_last_name TEXT NOT NULL
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS correspondence (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    date TEXT NOT NULL,
                    corr_type TEXT NOT NULL,
                    urgency TEXT NOT NULL,
                    incoming INTEGER,
                    outgoing INTEGER,
                    period TEXT NOT NULL,
                    duty_dus_id INTEGER,
                    FOREIGN KEY(duty_dus_id) REFERENCES duty_dus(id)
                )
            ''')
            connection.commit()
        except sqlite3.DatabaseError as e:
            QMessageBox.critical(self, "Ошибка базы данных", f"Не удалось создать таблицы: {str(e)}")
        finally:
            if connection:
                connection.close()

    def update_duty_dus_list(self):
        # Обновляем список дежурных по узлу связи в интерфейсе
        try:
            connection = sqlite3.connect(self.db_path)
            cursor = connection.cursor()
            cursor.execute('SELECT id, rank, first_name, last_name, last_last_name FROM duty_dus')
            duty_dus_list = cursor.fetchall()
        except sqlite3.DatabaseError as e:
            QMessageBox.critical(self, "Ошибка базы данных", f"Не удалось получить список дежурных: {str(e)}")
        else:
            self.duty_dus_input.clear()
            self.delete_duty_dus_input.clear()
            for duty_dus in duty_dus_list:
                duty_dus_id = duty_dus[0]
                rank = duty_dus[1]
                first_name = duty_dus[2]
                last_name = duty_dus[3]
                last_last_name = duty_dus[4]
                full_name = f"{rank} {first_name} {last_name} {last_last_name}"
                self.duty_dus_input.addItem(full_name, duty_dus_id)
                self.delete_duty_dus_input.addItem(full_name, duty_dus_id)
        finally:
            if connection:
                connection.close()

    def add_duty_dus(self):
        # Получаем данные о новом дежурном офицере
        rank = self.rank_input.text()
        first_name = self.first_name_input.text()
        last_name = self.last_name_input.text()
        last_last_name = self.last_last_name_input.text()

        # Проверка на пустые поля
        if not rank.strip() or not first_name.strip() or not last_name.strip() or not last_last_name.strip():
            QMessageBox.warning(self, "Ошибка ввода", "Все поля должны быть заполнены!")
            return

        try:
            # Подключаемся к базе данных и выполняем запрос
            connection = sqlite3.connect(self.db_path)
            cursor = connection.cursor()
            cursor.execute('''
                INSERT INTO duty_dus (rank, first_name, last_name, last_last_name)
                VALUES (?, ?, ?, ?)
            ''', (rank, first_name, last_name, last_last_name))

            # Подтверждаем изменения
            connection.commit()

            # Обновляем список дежурных офицеров
            self.update_duty_dus_list()

            # Очищаем поля ввода
            self.rank_input.clear()
            self.first_name_input.clear()
            self.last_name_input.clear()
            self.last_last_name_input.clear()

            # Сообщаем об успехе пользователю
            QMessageBox.information(self, "Успех", "Дежурный офицер добавлен.")

        except sqlite3.DatabaseError as e:
            # Обработка ошибки базы данных
            QMessageBox.critical(self, "Ошибка базы данных", f"Не удалось добавить дежурного офицера: {str(e)}")

        finally:
            # Закрываем соединение с базой данных
            if connection:
                connection.close()

    def add_correspondence(self):
        # Получаем данные из полей
        date = self.date_input.date().toString("yyyy-MM-dd")
        corr_type = self.corr_type_input.currentText()
        urgency = self.urgency_input.currentText()
        incoming = self.incoming_input.text()
        outgoing = self.outgoing_input.text()
        period = self.period_input.currentText()
        duty_dus_id = self.duty_dus_input.currentData()

        if not all([date, corr_type, urgency, incoming, outgoing, period, duty_dus_id]):
            QMessageBox.warning(self, "Ошибка ввода", "Все поля должны быть заполнены!")
            return
        try:
            # Подключаемся к базе данных и выполняем запрос с использованием контекстного менеджера
            with sqlite3.connect(self.db_path) as connection:
                cursor = connection.cursor()

                cursor.execute('''
                    INSERT INTO correspondence (date, corr_type, urgency, incoming, outgoing, period, duty_dus_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (date, corr_type, urgency, incoming, outgoing, period, duty_dus_id))

            # Очищаем поля ввода после успешного добавления
            self.incoming_input.clear()
            self.outgoing_input.clear()

            # Отправляем сообщение об успешном добавлении
            QMessageBox.information(self, "Успех", "Корреспонденция добавлена.")

        except sqlite3.DatabaseError as e:
            # Обработка ошибки базы данных
            print(f"Ошибка базы данных: {e}")
            QMessageBox.critical(self, "Ошибка базы данных", f"Не удалось добавить нагрузку: {str(e)}")

        # Контекстный менеджер автоматически закроет соединение с базой данных

    def delete_duty_dus(self):
        # Получаем ID дежурного офицера для удаления
        duty_dus_id = self.delete_duty_dus_input.currentData()

        # Отображаем диалог подтверждения
        reply = QMessageBox.question(
            self,
            "Подтверждение удаления",
            "Вы уверены, что хотите удалить этого дежурного по узлу связи?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        # Проверяем выбор пользователя
        if reply == QMessageBox.Yes:
            try:
                # Удаляем дежурного офицера из базы данных
                connection = sqlite3.connect(self.db_path)
                cursor = connection.cursor()
                cursor.execute('DELETE FROM duty_dus WHERE id = ?', (duty_dus_id,))
                connection.commit()
                QMessageBox.information(self, "Успех", "Дежурный по узлу связи удален.")
            except sqlite3.DatabaseError as e:
                QMessageBox.critical(self, "Ошибка базы данных",
                                     f"Не удалось удалить дежурного по узлу связи: {str(e)}")
            finally:
                if connection:
                    connection.close()

        # Обновляем список дежурных офицеров
        self.update_duty_dus_list()

    def search_correspondence(self):
        # Получаем диапазон дат для поиска
        start_date = self.start_date_input.date().toString("yyyy-MM-dd")
        end_date = self.end_date_input.date().toString("yyyy-MM-dd")

        try:
            # Выполняем запрос к базе данных для поиска корреспонденции
            connection = sqlite3.connect(self.db_path)
            cursor = connection.cursor()
            cursor.execute('''
                SELECT c.date, c.corr_type, c.urgency, c.incoming, c.outgoing, c.period,
                       d.rank, d.first_name, d.last_name, d.last_last_name
                FROM correspondence c
                JOIN duty_dus d ON c.duty_dus_id = d.id
                WHERE date(c.date) BETWEEN date(?) AND date(?)
            ''', (start_date, end_date))
            results = cursor.fetchall()
        except sqlite3.DatabaseError as e:
            QMessageBox.critical(self, "Ошибка базы данных", f"Не удалось выполнить поиск: {str(e)}")
        else:
            # Очистка и обновление таблицы результатов поиска
            self.search_results_table.clearContents()
            self.search_results_table.setRowCount(len(results))

            for i, row in enumerate(results):
                # Разбираем данные о дежурном
                rank = row[6]  # Звание
                first_name = row[7]  # Имя
                last_name = row[8]  # Фамилия
                last_last_name = row[9]  # Отчество

                # Формируем строку в нужном формате: "rank last_name initials"
                # Инициал имени
                initial_first_name = first_name[
                    0].upper()  # Берем первую букву имени и преобразуем ее в верхний регистр
                # Инициал отчества
                initial_last_last_name = last_last_name[
                    0].upper() if last_last_name else ''  # Берем первую букву отчества

                # Формируем строку дежурного по узлу связи
                duty_dus_str = f"{rank} {last_name} {initial_first_name}.{initial_last_last_name}."

                # Добавляем данные в таблицу результатов
                for j, value in enumerate(row[:6]):
                    item = QTableWidgetItem(str(value))
                    self.search_results_table.setItem(i, j, item)

                # Добавляем строку duty_dus_str в последний столбец таблицы результатов поиска
                item = QTableWidgetItem(duty_dus_str)
                self.search_results_table.setItem(i, 6, item)

        finally:
            # Закрываем соединение с базой данных
            if connection:
                connection.close()

    def export_to_pdf(self):
        # Экспорт результатов поиска в PDF-файл
        filename, _ = QFileDialog.getSaveFileName(self, "Сохранить PDF", "", "PDF Files (*.pdf)")

        if filename:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            # Заголовок
            pdf.set_font("Arial", style="B", size=14)
            pdf.cell(0, 10, "Результаты поиска корреспонденции", ln=True, align="C")
            pdf.set_font("Arial", size=12)

            # Заголовки столбцов
            headers = [
                "Дата", "Тип", "Срочность", "Входящая",
                "Исходящая", "Период", "ДУС"
            ]
            for header in headers:
                pdf.cell(30, 10, header, border=1)

            pdf.ln()

            # Добавляем данные в PDF-файл
            for i in range(self.search_results_table.rowCount()):
                for j in range(self.search_results_table.columnCount()):
                    cell_value = self.search_results_table.item(i, j).text()
                    pdf.cell(30, 10, cell_value, border=1)
                pdf.ln()

            pdf.output(filename)
            QMessageBox.information(self, "Успех", "Результаты поиска экспортированы в PDF")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    db_path = "correspondence.db"  # Укажите путь к вашей базе данных
    window = CorrespondenceApp(db_path)
    window.show()
    sys.exit(app.exec_())
