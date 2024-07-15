import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, 
    QPushButton, QFileDialog, QComboBox, QMessageBox
)
from functions import convert_excel, convert_json_to_csv, convert_csv_to_excel

class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('EFC')
        self.setGeometry(100, 100, 400, 300)

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
        """)

        layout = QVBoxLayout()

        self.input_label = QLabel('Input File:', self)
        layout.addWidget(self.input_label)

        self.input_line_edit = QLineEdit(self)
        layout.addWidget(self.input_line_edit)

        self.input_button = QPushButton('Browse', self)
        self.input_button.clicked.connect(self.browse_input_file)
        layout.addWidget(self.input_button)

        self.output_label = QLabel('Output File:', self)
        layout.addWidget(self.output_label)

        self.output_line_edit = QLineEdit(self)
        layout.addWidget(self.output_line_edit)

        self.output_button = QPushButton('Browse', self)
        self.output_button.clicked.connect(self.browse_output_file)
        layout.addWidget(self.output_button)

        self.type_label = QLabel('Conversion Type:', self)
        layout.addWidget(self.type_label)

        self.type_combo = QComboBox(self)
        self.type_combo.addItems(['Excel to CSV', 'CSV to Excel', 'JSON to CSV'])
        layout.addWidget(self.type_combo)

        self.convert_button = QPushButton('Convert', self)
        self.convert_button.clicked.connect(self.convert_file)
        layout.addWidget(self.convert_button)

        self.setLayout(layout)

    def browse_input_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Input File", "", 
                                                  "All Files (*);;Excel Files (*.xlsx);;CSV Files (*.csv);;JSON Files (*.json)", options=options)
        if file_name:
            self.input_line_edit.setText(file_name)

    def browse_output_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Select Output File", "", 
                                                  "All Files (*);;CSV Files (*.csv);;Excel Files (*.xlsx)", options=options)
        if file_name:
            self.output_line_edit.setText(file_name)

    def convert_file(self):
        input_file = self.input_line_edit.text()
        output_file = self.output_line_edit.text()
        conversion_type = self.type_combo.currentText()

        if conversion_type == 'Excel to CSV':
            convert_excel(input_file, output_file)
        elif conversion_type == 'CSV to Excel':
            convert_csv_to_excel(input_file, output_file)
        elif conversion_type == 'JSON to CSV':
            convert_json_to_csv(input_file, output_file)
        else:
            QMessageBox.warning(self, "Conversion Type Error", "Invalid conversion type selected.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ConverterApp()
    ex.show()
    sys.exit(app.exec_())