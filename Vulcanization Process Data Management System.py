
import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QMessageBox, QTableWidget, QTableWidgetItem, QDialog
)
from PyQt5.QtGui import QFont, QColor
from PyQt5.QtCore import Qt
import pyodbc


conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\Ali_Shawrma\Diplom\Vulcanization.accdb;'
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()


def fetch_data(table_name):
    cursor.execute(f'SELECT * FROM {table_name}')
    return cursor.fetchall(), [desc[0] for desc in cursor.description]


def insert_data(table_name, columns, values):
    placeholders = ', '.join(['?' for _ in values])
    cursor.execute(f'INSERT INTO {table_name} ({columns}) VALUES ({placeholders})', values)
    conn.commit()

def update_data(table_name, columns, values, key_column, key_value):
    set_clause = ', '.join([f'{col}=?' for col in columns if col != key_column])  # Exclude primary key column from SET clause
    update_query = f'UPDATE {table_name} SET {set_clause} WHERE {key_column}=?'


    values_to_update = [value for col, value in zip(columns, values) if col != key_column]

    try:
        cursor.execute(update_query, values_to_update + [key_value])
        conn.commit()
        print("Данные успешно обновлены")
    except Exception as e:
        print(f"Exception in update_data: {e}")
        raise

def delete_data(table_name, key_column, key_value):
    cursor.execute(f'DELETE FROM {table_name} WHERE {key_column}=?', (key_value,))
    conn.commit()


class DataForm(QDialog):
    def __init__(self, table_name, columns, parent=None):
        super().__init__(parent)
        self.table_name = table_name
        self.columns = columns
        self.setWindowTitle('Добавить/Изменить запись')
        self.setGeometry(100, 100, 300, 400)
        self.setStyleSheet("background-color: #f0f0f0;")
        self.layout = QVBoxLayout()

        self.inputs = {}
        for column in columns:
            label = QLabel(f'{column}:')
            input_field = QLineEdit(self)
            label.setFont(QFont("Arial", 12))
            input_field.setFont(QFont("Arial", 12))
            self.layout.addWidget(label)
            self.layout.addWidget(input_field)
            self.inputs[column] = input_field

        self.saveButton = QPushButton('Сохранить', self)
        self.saveButton.setFont(QFont("Arial", 12, QFont.Bold))
        self.saveButton.setStyleSheet("background-color: #4CAF50; color: white;")
        self.saveButton.clicked.connect(self.save_data)
        self.layout.addWidget(self.saveButton)

        self.setLayout(self.layout)

    def save_data(self):
        values = [self.inputs[col].text() for col in self.columns]
        try:
            insert_data(self.table_name, ', '.join(self.columns), values)
            QMessageBox.information(self, 'Успех', 'Запись успешно добавлена!')
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', str(e))

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.current_table = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Управление базой данных вулканизации')
        self.setGeometry(100, 100, 800, 600)
        self.setStyleSheet("background-color: #e0f7fa;")

        self.layout = QVBoxLayout()

        self.tableButtonsLayout = QHBoxLayout()

        button_style = """
        QPushButton {
            background-color: #03a9f4;
            color: white;
            font-size: 14px;
            font-weight: bold;
            padding: 10px;
            border-radius: 5px;
        }
        QPushButton:hover {
            background-color: #0288d1;
        }
        """

        self.samplesButton = QPushButton('Наименование экспериментальных образцов', self)
        self.samplesButton.setStyleSheet(button_style)
        self.samplesButton.clicked.connect(lambda: self.load_table('наименование_экспериментальных_образцов'))
        self.tableButtonsLayout.addWidget(self.samplesButton)

        self.componentButton = QPushButton('Компонентный состав образцов', self)
        self.componentButton.setStyleSheet(button_style)
        self.componentButton.clicked.connect(lambda: self.load_table('компонентный_состав_образцов'))
        self.tableButtonsLayout.addWidget(self.componentButton)

        self.tempTimeButton = QPushButton('Условия проведения температурно временного эксперимента', self)
        self.tempTimeButton.setStyleSheet(button_style)
        self.tempTimeButton.clicked.connect(lambda: self.load_table('условия_проведения_температурно_временного_эксперимента'))
        self.tableButtonsLayout.addWidget(self.tempTimeButton)

        self.rheoButton = QPushButton('Условия проведения реометрического эксперимента', self)
        self.rheoButton.setStyleSheet(button_style)
        self.rheoButton.clicked.connect(lambda: self.load_table('условия_проведения_реометрического_эксперимента'))
        self.tableButtonsLayout.addWidget(self.rheoButton)

        self.layout.addLayout(self.tableButtonsLayout)

        self.tableWidget = QTableWidget()
        self.tableWidget.setStyleSheet("""
        QTableWidget {
            background-color: #ffffff;
            font-size: 14px;
        }
        QHeaderView::section {
            background-color: #0288d1;
            color: white;
            font-size: 14px;
            font-weight: bold;
        }
        """)
        self.layout.addWidget(self.tableWidget)

        self.buttonLayout = QHBoxLayout()

        self.addButton = QPushButton('Добавить')
        self.addButton.setStyleSheet(button_style)
        self.addButton.clicked.connect(self.add_record)
        self.buttonLayout.addWidget(self.addButton)

        self.modifyButton = QPushButton('Изменить')
        self.modifyButton.setStyleSheet(button_style)
        self.modifyButton.clicked.connect(self.modify_record)
        self.buttonLayout.addWidget(self.modifyButton)

        self.deleteButton = QPushButton('Удалить')
        self.deleteButton.setStyleSheet(button_style)
        self.deleteButton.clicked.connect(self.delete_record)
        self.buttonLayout.addWidget(self.deleteButton)

        self.layout.addLayout(self.buttonLayout)

        self.setLayout(self.layout)
        self.show()

    def load_table(self, table_name):
        self.current_table = table_name
        data, columns = fetch_data(table_name)
        self.tableWidget.setRowCount(0)
        self.tableWidget.setColumnCount(len(columns))
        self.tableWidget.setHorizontalHeaderLabels(columns)

        for row_number, row_data in enumerate(data):
            self.tableWidget.insertRow(row_number)
            for column_number, data in enumerate(row_data):
                self.tableWidget.setItem(row_number, column_number, QTableWidgetItem(str(data)))

    def add_record(self):
        if not self.current_table:
            QMessageBox.warning(self, 'Предупреждение', 'Пожалуйста, сначала выберите таблицу')
            return

        data, columns = fetch_data(self.current_table)
        dialog = DataForm(self.current_table, columns, self)
        if dialog.exec_():
            self.load_table(self.current_table)

    def modify_record(self):
        if not self.current_table:
            QMessageBox.warning(self, 'Предупреждение', 'Пожалуйста, сначала выберите таблицу')
            return

        selected_row = self.tableWidget.currentRow()
        if selected_row < 0:
            QMessageBox.warning(self, 'Предупреждение', 'Пожалуйста, выберите запись для изменения')
            return

        data, columns = fetch_data(self.current_table)
        key_value = self.tableWidget.item(selected_row, 0).text()

        dialog = DataForm(self.current_table, columns, self)
        for col in range(len(columns)):
            dialog.inputs[columns[col]].setText(self.tableWidget.item(selected_row, col).text())

        if dialog.exec_():
            values = [dialog.inputs[col].text() for col in columns]

            try:
                print(f"Updating table: {self.current_table}")
                print(f"Columns: {columns}")
                print(f"Values: {values}")
                print(f"Key column: {columns[0]}")
                print(f"Key value: {key_value}")

                update_data(self.current_table, columns, values, columns[0], key_value)
                QMessageBox.information(self, 'Успех', 'Запись успешно изменена!')
                self.load_table(self.current_table)
            except Exception as e:
                print(f"Exception: {e}")
                QMessageBox.critical(self, 'Ошибка', str(e))

    def delete_record(self):
        if not self.current_table:
            QMessageBox.warning(self, 'Предупреждение', 'Пожалуйста, сначала выберите таблицу')
            return

        selected_row = self.tableWidget.currentRow()
        if selected_row < 0:
            QMessageBox.warning(self, 'Предупреждение', 'Пожалуйста, выберите запись для удаления')
            return

        key_value = self.tableWidget.item(selected_row, 0).text()

        try:
            delete_data(self.current_table, self.tableWidget.horizontalHeaderItem(0).text(), key_value)
            QMessageBox.information(self, 'Успех', 'Запись успешно удалена!')
            self.load_table(self.current_table)
        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', str(e))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
