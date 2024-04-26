import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QLineEdit, QPushButton, QComboBox, QTableWidget, QTableWidgetItem,
                             QDateEdit, QTabWidget, QMessageBox, QFileDialog, QFormLayout, QCheckBox)
from PyQt5.QtCore import QDate,Qt
import sqlite3
from contextlib import contextmanager
from fpdf import FPDF
from datetime import datetime

@contextmanager
def db_connection(db_path):
    try:
        connection = sqlite3.connect(db_path)
        yield connection
    finally:
        connection.close()

class DatabaseManager:
    def __init__(self, db_path):
        self.db_path = db_path

    def create_tables(self):
        with db_connection(self.db_path) as connection:
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

    def add_duty_dus(self, rank, first_name, last_name, last_last_name):
        with db_connection(self.db_path) as connection:
            cursor = connection.cursor()
            cursor.execute('''
                INSERT INTO duty_dus (rank, first_name, last_name, last_last_name)
                VALUES (?, ?, ?, ?)
            ''', (rank, first_name, last_name, last_last_name))

    def add_correspondence(self, date, corr_type, urgency, incoming, outgoing, period, duty_dus_id):
        with db_connection(self.db_path) as connection:
            cursor = connection.cursor()
            cursor.execute('''
                INSERT INTO correspondence (date, corr_type, urgency, incoming, outgoing, period, duty_dus_id)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (date, corr_type, urgency, incoming, outgoing, period, duty_dus_id))

    def update_duty_dus_list(self):
        with db_connection(self.db_path) as connection:
            cursor = connection.cursor()
            cursor.execute('SELECT id, rank, first_name, last_name, last_last_name FROM duty_dus')
            return cursor.fetchall()

    def search_correspondence(self, start_date, end_date):
        with db_connection(self.db_path) as connection:
            cursor = connection.cursor()
            cursor.execute('''
                SELECT c.date, c.corr_type, c.urgency, c.incoming, c.outgoing, c.period,
                       d.rank, d.first_name, d.last_name, d.last_last_name
                FROM correspondence c
                JOIN duty_dus d ON c.duty_dus_id = d.id
                WHERE date(c.date) BETWEEN date(?) AND date(?)
            ''', (start_date, end_date))
            return cursor.fetchall()

    def get_correspondence_count_by_type(self, start_date, end_date):
        with db_connection(self.db_path) as connection:
            cursor = connection.cursor()
            cursor.execute('''
                SELECT corr_type, urgency, SUM(incoming) AS incoming_total, SUM(outgoing) AS outgoing_total
                FROM correspondence
                WHERE date(date) BETWEEN date(?) AND date(?)
                GROUP BY corr_type, urgency
            ''', (start_date, end_date))
            return cursor.fetchall()

class CorrespondenceApp(QWidget):
    def __init__(self, db_path):
        super().__init__()
        self.db_path = db_path
        self.db_manager = DatabaseManager(db_path)
        self.init_ui()
        self.create_tables()
        self.update_export_duty_dus_list()

    def init_ui(self):
        self.setWindowTitle("Система учета корреспонденции")
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget)

        self.create_add_correspondence_tab(tab_widget)
        self.create_manage_duty_dus_tab(tab_widget)
        self.create_search_correspondence_tab(tab_widget)

        self.setStyleSheet(self.get_stylesheet())

    def get_stylesheet(self):
        stylesheet = """
            QWidget {
                font-family: Arial;
                font-size: 11pt;
            }

            QTabWidget {
                border: none;
            }

            QTabWidget::pane {
                border: none;
            }

            QTabBar::tab {
                margin-left: 5px;
                margin-right: 5px;
                padding: 5px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                border: 1px solid #ccc;
                color: #555;
            }

            QTabBar::tab:selected {
                background-color: #f0f0f0;
                border-bottom-color: #f0f0f0;
                color: #333;
            }

            QPushButton {
                border: 1px solid #555;
                border-radius: 5px;
                padding: 5px;
            }

            QPushButton:hover {
                background-color: #f4f4f4;
            }

            QPushButton:pressed {
                background-color: #eaeaea;
            }

            QLineEdit, QDateEdit, QComboBox {
                border:hover {
                    border: 1px solid #aaa;
                }
                border: 1px solid #ccc;
                border-radius: 3px;
                padding: 3px;
            }

            QTableWidget {
                border: 1px solid #ccc;
            }

            QTableWidget::item {
                border-bottom: 1px solid #ddd;
                border-right: 1px solid #ddd;
                padding: 5px;
            }

            QHeaderView::section {
                background-color: #f0f0f0;
                border: 1px solid #ccc;
                padding: 3px;
            }
        """
        return stylesheet

    def create_add_correspondence_tab(self, tab_widget):
        add_correspondence_tab = QWidget()
        layout = QFormLayout()
        add_correspondence_tab.setLayout(layout)

        self.date_input = QDateEdit()
        self.date_input.setDate(QDate.currentDate())
        self.date_input.setCalendarPopup(True)

        layout.addRow(QLabel("Дата:"), self.date_input)

        self.corr_type_input = QComboBox()
        self.corr_type_input.addItems(["ТЛФ", "ТЛГ", "ЗС СПД"])

        layout.addRow(QLabel("Тип информации:"), self.corr_type_input)

        self.urgency_input = QComboBox()
        self.urgency_input.addItems([
            "ОБК", "ДСП", "Секр.", "Срч.",
            "Смл", "Ркт", "Воздух", "Мнлт",
            "Секр. срч.", "Секр. смл.", "Секр. мнлт.",
            "СС", "ОВ"
        ])

        layout.addRow(QLabel("Срочность:"), self.urgency_input)

        self.incoming_input = QLineEdit()
        layout.addRow(QLabel("Кол-во входящих:"), self.incoming_input)

        self.outgoing_input = QLineEdit()
        layout.addRow(QLabel("Кол-во исходящих:"), self.outgoing_input)

        self.period_input = QComboBox()  # Создайте объект QComboBox перед использованием
        self.period_input.addItems([
            "С 10:00 по 22:00",
            "С 22:00 по 10:00"
        ])

        layout.addRow(QLabel("Период дежурства:"), self.period_input)

        self.duty_dus_input = QComboBox()
        layout.addRow(QLabel("ДУС:"), self.duty_dus_input)


        add_button = QPushButton("Добавить")
        add_button.clicked.connect(self.add_correspondence)
        layout.addRow(add_button)

        tab_widget.addTab(add_correspondence_tab, "Добавлении нагрузки")
    def create_manage_duty_dus_tab(self, tab_widget):
        manage_duty_dus_tab = QWidget()
        layout = QVBoxLayout()
        manage_duty_dus_tab.setLayout(layout)

        add_layout = QVBoxLayout()
        add_layout.addWidget(QLabel("Добавить дежурного по узлу связи"))
        layout.addLayout(add_layout)

        rank_layout = QHBoxLayout()
        rank_layout.addWidget(QLabel("Звание:"))
        self.rank_input = QComboBox()
        self.rank_input.addItems(["ефр.", "пр-к", "ст. л-т"])
        rank_layout.addWidget(self.rank_input)
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

        delete_layout = QVBoxLayout()
        delete_layout.addWidget(QLabel("Удаление дежурного по узлу связи"))
        layout.addLayout(delete_layout)

        self.delete_duty_dus_input = QComboBox()
        delete_layout.addWidget(self.delete_duty_dus_input)

        delete_button = QPushButton("Удалить")
        delete_button.clicked.connect(self.delete_duty_dus)
        delete_layout.addWidget(delete_button)

        tab_widget.addTab(manage_duty_dus_tab, "Управление ДУС")

    def create_search_correspondence_tab(self, tab_widget):
        search_correspondence_tab = QWidget()
        layout = QVBoxLayout()
        search_correspondence_tab.setLayout(layout)

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

        self.search_results_table = QTableWidget()
        self.search_results_table.setColumnCount(7)
        self.search_results_table.setHorizontalHeaderLabels([
            "Дата", "Тип", "Срочность", "Входящая", "Исходящая", "Период", "ДУС"
        ])
        layout.addWidget(self.search_results_table)

        self.category_totals_table = QTableWidget()
        self.category_totals_table.setColumnCount(4)
        self.category_totals_table.setHorizontalHeaderLabels([
            "Тип", "Срочность", "Входящая", "Исходящая"
        ])
        layout.addWidget(self.category_totals_table)

        self.export_duty_dus_input = QComboBox()
        self.update_export_duty_dus_list()
        layout.addWidget(QLabel("Экспорт ДУС:"))
        layout.addWidget(self.export_duty_dus_input)

        self.export_search_results_checkbox = QCheckBox("Отобразить результаты всего поиска")
        self.export_search_results_checkbox.setChecked(True)
        self.export_search_results_checkbox.stateChanged.connect(self.toggle_search_results_export)
        layout.addWidget(self.export_search_results_checkbox)

        export_button = QPushButton("Экспортировать в PDF")
        export_button.clicked.connect(self.export_to_pdf)
        layout.addWidget(export_button)

        tab_widget.addTab(search_correspondence_tab, "Статистика нагрузки")

    def toggle_search_results_export(self, state):
        if state == Qt.Checked:
            self.search_results_table.show()
        else:
            self.search_results_table.hide()

    def create_tables(self):
        self.db_manager.create_tables()

    def update_export_duty_dus_list(self):
        duty_dus_list = self.db_manager.update_duty_dus_list()
        self.export_duty_dus_input.clear()
        for duty_dus in duty_dus_list:
            rank = duty_dus[1]
            first_name = duty_dus[2]
            last_name = duty_dus[3]
            last_last_name = duty_dus[4]
            full_name = f"{rank} {first_name} {last_name} {last_last_name}"
            self.export_duty_dus_input.addItem(full_name)

    def add_duty_dus(self):
        global connection
        rank = self.rank_input.currentText()
        first_name = self.first_name_input.text()
        last_name = self.last_name_input.text()
        last_last_name = self.last_last_name_input.text()

        if not rank.strip() or not first_name.strip() or not last_name.strip() or not last_last_name.strip():
            QMessageBox.warning(self, "Ошибка ввода", "Все поля должны быть заполнены!")
            return

        try:
            connection = sqlite3.connect(self.db_path)
            cursor = connection.cursor()
            cursor.execute('''
                INSERT INTO duty_dus (rank, first_name, last_name, last_last_name)
                VALUES (?, ?, ?, ?)
            ''', (rank, first_name, last_name, last_last_name))
            connection.commit()
            self.update_duty_dus_list()
            self.first_name_input.clear()
            self.last_name_input.clear()
            self.last_last_name_input.clear()
            QMessageBox.information(self, "Успех", "Дежурный по узлу связи добавлен.")
        except sqlite3.DatabaseError as e:
            QMessageBox.critical(self, "Ошибка базы данных", f"Не удалось добавить дежурного по узлу связи: {str(e)}")
        finally:
            if connection:
                connection.close()

    def add_correspondence(self):
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
            with sqlite3.connect(self.db_path) as connection:
                cursor = connection.cursor()
                cursor.execute('''
                    INSERT INTO correspondence (date, corr_type, urgency, incoming, outgoing, period, duty_dus_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (date, corr_type, urgency, incoming, outgoing, period, duty_dus_id))
            self.incoming_input.clear()
            self.outgoing_input.clear()
            QMessageBox.information(self, "Успех", "Корреспонденция добавлена.")
        except sqlite3.DatabaseError as e:
            QMessageBox.critical(self, "Ошибка базы данных", f"Не удалось добавить нагрузку: {str(e)}")

    def delete_duty_dus(self):
        duty_dus_id = self.delete_duty_dus_input.currentData()

        reply = QMessageBox.question(
            self,
            "Подтверждение удаления",
            "Вы уверены, что хотите удалить этого дежурного по узлу связи?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            try:
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
            self.update_duty_dus_list()

    def search_correspondence(self):
        start_date = self.start_date_input.date().toString("yyyy-MM-dd")
        end_date = self.end_date_input.date().toString("yyyy-MM-dd")

        results = self.db_manager.search_correspondence(start_date, end_date)

        self.search_results_table.clearContents()
        self.search_results_table.setRowCount(len(results))

        for i, row in enumerate(results):
            rank = row[6]
            first_name = row[7]
            last_name = row[8]
            last_last_name = row[9]

            initial_first_name = first_name[0].upper()
            initial_last_last_name = last_last_name[0].upper() if last_last_name else ''

            duty_dus_str = f"{rank} {last_name} {initial_first_name}.{initial_last_last_name}."

            for j, value in enumerate(row[:6]):
                item = QTableWidgetItem(str(value))
                self.search_results_table.setItem(i, j, item)

            item = QTableWidgetItem(duty_dus_str)
            self.search_results_table.setItem(i, 6, item)

        # Calculate and display category totals
        self.category_totals_table.clearContents()
        self.category_totals_table.setRowCount(0)

        category_totals = self.db_manager.get_correspondence_count_by_type(start_date, end_date)

        self.category_totals_table.setRowCount(len(category_totals))

        for i, row in enumerate(category_totals):
            corr_type = row[0]
            urgency = row[1]
            incoming_total = row[2]
            outgoing_total = row[3]

            item = QTableWidgetItem(corr_type)
            self.category_totals_table.setItem(i, 0, item)

            item = QTableWidgetItem(urgency)
            self.category_totals_table.setItem(i, 1, item)

            item = QTableWidgetItem(str(incoming_total))
            self.category_totals_table.setItem(i, 2, item)

            item = QTableWidgetItem(str(outgoing_total))
            self.category_totals_table.setItem(i, 3, item)

    def export_to_pdf(self):
        current_datetime = datetime.now()
        formatted_datetime = current_datetime.strftime('%d-%m-%Y_%H-%M')
        filename = f"{formatted_datetime}_report.pdf"
        full_filename, _ = QFileDialog.getSaveFileName(self, "Сохранить PDF", filename, "PDF Files (*.pdf)")

        if full_filename:
            pdf = FPDF(orientation="P", unit="mm", format="A4")
            font_path_regular = 'DejaVuSansCondensed.ttf'
            font_path_bold = 'DejaVuSansCondensed-Bold.ttf'
            font_path_italic = 'DejaVuSerifCondensed-Italic.ttf'
            pdf.add_font('DejaVu', '', font_path_regular, uni=True)
            pdf.add_font('DejaVu', 'B', font_path_bold, uni=True)
            pdf.add_font('DejaVu', 'I', font_path_italic, uni=True)
            pdf.add_page()
            pdf.set_font("DejaVu", size=8)

            if self.export_search_results_checkbox.isChecked():
                pdf.set_font("DejaVu", style="B", size=10)
                pdf.cell(0, 10, "Результаты поиска корреспонденции", ln=True, align="C")
                pdf.set_font("DejaVu", size=8)
                headers = ["Дата", "Тип", "Срочность", "Входящая", "Исходящая", "Период", "ДУС"]
                column_widths = [25, 25, 30, 20, 20, 35, 40]
                for i, header in enumerate(headers):
                    pdf.cell(column_widths[i], 10, header, border=1, align="C")
                pdf.ln()
                for i in range(self.search_results_table.rowCount()):
                    for j in range(self.search_results_table.columnCount()):
                        cell_value = self.search_results_table.item(i, j).text()
                        pdf.cell(column_widths[j], 10, cell_value, border=1, align="C")
                    pdf.ln()

            pdf.ln(20)
            pdf.set_font("DejaVu", style="B", size=10)
            pdf.cell(0, 10, "Итоги по категориям", ln=True, align="C")
            pdf.set_font("DejaVu", size=8, align="C")
            headers = ["Тип", "Срочность", "Входящая", "Исходящая"]
            column_widths = [40, 30, 20, 20]
            for i, header in enumerate(headers):
                pdf.cell(column_widths[i], 10, header, border=1, align="C")
            pdf.ln()
            for i in range(self.category_totals_table.rowCount()):
                for j in range(self.category_totals_table.columnCount()):
                    cell_value = self.category_totals_table.item(i, j).text()
                    pdf.cell(column_widths[j], 10, cell_value, border=1, align="C")
                pdf.ln()

            pdf.ln(20)
            pdf.set_font("DejaVu", style="I", size=8)
            pdf.cell(0, 10, f"Выгрузка выполнено: {formatted_datetime}", ln=True, align="C")
            pdf.cell(0, 10, "Выгрузил:", ln=True, align="C")
            selected_dus = self.export_duty_dus_input.currentText()
            pdf.cell(0, 10, selected_dus, ln=True, align="C")

            pdf.output(full_filename)
            QMessageBox.information(self, "Успех", "Результаты поиска экспортированы в PDF")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    db_path = "correspondence.db"
    window = CorrespondenceApp(db_path)
    window.show()
    sys.exit(app.exec_())
