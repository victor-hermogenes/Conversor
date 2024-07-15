import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, 
    QPushButton, QFileDialog, QComboBox, QMessageBox, QCheckBox, 
    QScrollArea, QFormLayout, QTableWidget, QTableWidgetItem, QHBoxLayout
)
from functions import convert_excel, convert_json_to_csv, convert_csv_to_excel

class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('EFC')
        self.setGeometry(100, 100, 800, 500)

        self.setStyleSheet("""
            QWidget {
                background-color: #2E2E2E;
                color: #FFFFFF;
            }
            QLabel {
                font-size: 14px;
                margin-bottom: 5px;
                color: #FFFFFF;
            }
            QLineEdit {
                font-size: 14px;
                padding: 5px;
                background-color: #3E3E3E;
                color: #FFFFFF;
                border: 1px solid #5A5A5A;
                border-radius: 3px;
            }
            QComboBox {
                font-size: 14px;
                padding: 5px;
                background-color: #3E3E3E;
                color: #FFFFFF;
                border: 1px solid #5A5A5A;
                border-radius: 3px;
            }
            QPushButton {
                font-size: 14px;
                padding: 10px;
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #388E3C;
            }
            QTableWidget {
                background-color: #3E3E3E;
                color: #FFFFFF;
                border: 1px solid #5A5A5A;
                gridline-color: #5A5A5A;
            }
            QHeaderView::section {
                background-color: #3E3E3E;
                color: #FFFFFF;
                border: 1px solid #5A5A5A;
            }
        """)

        main_layout = QHBoxLayout()
        self.setLayout(main_layout)

        left_layout = QVBoxLayout()
        main_layout.addLayout(left_layout)

        self.input_label = QLabel('Input File:', self)
        left_layout.addWidget(self.input_label)

        self.input_line_edit = QLineEdit(self)
        left_layout.addWidget(self.input_line_edit)

        self.input_button = QPushButton('Browse', self)
        self.input_button.clicked.connect(self.browse_input_file)
        left_layout.addWidget(self.input_button)

        self.output_label = QLabel('Output File:', self)
        left_layout.addWidget(self.output_label)

        self.output_line_edit = QLineEdit(self)
        left_layout.addWidget(self.output_line_edit)

        self.output_button = QPushButton('Browse', self)
        self.output_button.clicked.connect(self.browse_output_file)
        left_layout.addWidget(self.output_button)

        self.type_label = QLabel('Conversion Type:', self)
        left_layout.addWidget(self.type_label)

        self.type_combo = QComboBox(self)
        self.type_combo.addItems(['Excel to CSV', 'CSV to Excel', 'JSON to CSV'])
        left_layout.addWidget(self.type_combo)

        self.columns_label = QLabel('Select Columns:', self)
        left_layout.addWidget(self.columns_label)

        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget(self.scroll_area)
        self.scroll_layout = QFormLayout(self.scroll_content)
        self.scroll_area.setWidget(self.scroll_content)
        left_layout.addWidget(self.scroll_area)

        self.convert_button = QPushButton('Convert', self)
        self.convert_button.clicked.connect(self.convert_file)
        left_layout.addWidget(self.convert_button)

        # Table to display data
        self.table_widget = QTableWidget(self)
        main_layout.addWidget(self.table_widget)

    def browse_input_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Input File", "", 
                                                  "All Files (*);;Excel Files (*.xlsx);;CSV Files (*.csv);;JSON Files (*.json)", options=options)
        if file_name:
            self.input_line_edit.setText(file_name)
            self.update_columns(file_name)
            self.update_table_preview(file_name)

    def browse_output_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Select Output File", "", 
                                                  "All Files (*);;CSV Files (*.csv);;Excel Files (*.xlsx)", options=options)
        if file_name:
            self.output_line_edit.setText(file_name)

    def update_columns(self, file_name):
        self.clear_columns()
        if file_name.endswith('.csv'):
            import pandas as pd
            df = pd.read_csv(file_name, nrows=1)
            columns = df.columns.tolist()
        elif file_name.endswith('.xlsx'):
            import pandas as pd
            df = pd.read_excel(file_name, nrows=1)
            columns = df.columns.tolist()
        elif file_name.endswith('.json'):
            import json
            with open(file_name, 'r', encoding='utf-8') as f:
                first_line = f.readline().strip()
                first_record = json.loads(first_line)
                columns = first_record.keys()

        for column in columns:
            checkbox = QCheckBox(column, self)
            self.scroll_layout.addRow(checkbox)

    def clear_columns(self):
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def update_table_preview(self, file_name):
        if file_name.endswith('.csv'):
            import pandas as pd
            df = pd.read_csv(file_name, nrows=10)
        elif file_name.endswith('.xlsx'):
            import pandas as pd
            df = pd.read_excel(file_name, nrows=10)
        elif file_name.endswith('.json'):
            import json
            import pandas as pd
            data = []
            with open(file_name, 'r', encoding='utf-8') as f:
                for i, line in enumerate(f):
                    if i >= 10:
                        break
                    data.append(json.loads(line.strip()))
            df = pd.json_normalize(data)

        self.table_widget.setColumnCount(len(df.columns))
        self.table_widget.setRowCount(len(df.index))
        self.table_widget.setHorizontalHeaderLabels(df.columns)

        for row in df.iterrows():
            for col_index, value in enumerate(row[1]):
                self.table_widget.setItem(row[0], col_index, QTableWidgetItem(str(value)))

    def convert_file(self):
        input_file = self.input_line_edit.text()
        output_file = self.output_line_edit.text()
        conversion_type = self.type_combo.currentText()

        selected_columns = [self.scroll_layout.itemAt(i).widget().text()
                            for i in range(self.scroll_layout.count())
                            if self.scroll_layout.itemAt(i).widget().isChecked()]

        if conversion_type == 'Excel to CSV':
            convert_excel(input_file, output_file, selected_columns)
        elif conversion_type == 'CSV to Excel':
            convert_csv_to_excel(input_file, output_file, selected_columns)
        elif conversion_type == 'JSON to CSV':
            convert_json_to_csv(input_file, output_file, selected_columns)
        else:
            QMessageBox.warning(self, "Conversion Type Error", "Invalid conversion type selected.")

if __name__ == '__main__':
    try:
        app = QApplication(sys.argv)
        ex = ConverterApp()
        ex.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"Error: {e}")