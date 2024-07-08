import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QComboBox, QMessageBox
from functions import convert_excel

class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('EFC')
        self.setGeometry(100, 100, 400, 250)

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

        self.output_label = QLabel('Output Folder:', self)
        layout.addWidget(self.output_label)

        self.output_line_edit = QLineEdit(self)
        layout.addWidget(self.output_line_edit)

        self.output_button = QPushButton('Browse', self)
        self.output_button.clicked.connect(self.browse_output_folder)
        layout.addWidget(self.output_button)

        self.format_label = QLabel('Output Format:', self)
        layout.addWidget(self.format_label)

        self.format_combo_box = QComboBox(self)
        self.format_combo_box.addItems(["xlsx", "xls", "csv"])
        layout.addWidget(self.format_combo_box)

        self.convert_button = QPushButton('Convert', self)
        self.convert_button.clicked.connect(self.convert)
        layout.addWidget(self.convert_button)

        self.setLayout(layout)

    def browse_input_file(self):
        options = QFileDialog.Options()
        file, _ = QFileDialog.getOpenFileName(self, "Select Input File", "", "Excel Files (*.xlsx *.xls)", options=options)
        if file:
            self.input_line_edit.setText(file)

    def browse_output_folder(self):
        options = QFileDialog.Options()
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder", options=options)
        if folder:
            self.output_line_edit.setText(folder)

    def convert(self):
        input_file = self.input_line_edit.text()
        output_folder = self.output_line_edit.text()
        output_format = self.format_combo_box.currentText()

        if input_file and output_folder:
            input_filename = os.path.basename(input_file)
            base_name = os.path.splitext(input_filename)[0]
            output_file = os.path.join(output_folder, f"{base_name}.{output_format}")

            convert_excel(input_file, output_file)

        QMessageBox.information(self, "Conversiom Complete", f"File converted successfully to {output_file}")


def main():
    app = QApplication(sys.argv)
    ex = ConverterApp()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()