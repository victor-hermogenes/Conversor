import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, 
    QPushButton, QFileDialog, QComboBox, QMessageBox, QCheckBox, 
    QScrollArea, QFormLayout, QTableWidget, QTableWidgetItem, QHBoxLayout, QTabWidget, QToolButton, QStyle, QTabBar
)
from PyQt5.QtCore import Qt, QSize
from functions import convert_excel, convert_json_to_csv, convert_csv_to_excel

class FileConfig(QWidget):
    def __init__(self, file_path, file_name, close_callback, parent):
        super().__init__(parent)
        self.file_path = os.path.normpath(file_path)
        self.file_name = file_name
        self.close_callback = close_callback
        self.parent = parent
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        self.setLayout(layout)

        self.type_label = QLabel(f'Conversion Type for {self.file_name}:', self)
        layout.addWidget(self.type_label)

        self.type_combo = QComboBox(self)
        file_extension = os.path.splitext(self.file_name)[1].lower()
        if file_extension == '.xlsx':
            self.type_combo.addItems(['Excel to CSV'])
        elif file_extension == '.csv':
            self.type_combo.addItems(['CSV to Excel'])
        elif file_extension == '.json':
            self.type_combo.addItems(['JSON to CSV'])
        layout.addWidget(self.type_combo)

        self.columns_label = QLabel('Select Columns:', self)
        layout.addWidget(self.columns_label)

        self.select_all_checkbox = QCheckBox('Select All', self)
        self.select_all_checkbox.stateChanged.connect(self.toggle_select_all)
        layout.addWidget(self.select_all_checkbox)

        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget(self.scroll_area)
        self.scroll_layout = QFormLayout(self.scroll_content)
        self.scroll_area.setWidget(self.scroll_content)
        layout.addWidget(self.scroll_area)

        self.fragment_checkbox = QCheckBox('Fragmentate', self)
        self.fragment_checkbox.stateChanged.connect(self.toggle_fragmentation)
        layout.addWidget(self.fragment_checkbox)

        self.fragment_limit_label = QLabel('Fragment Size Limit (MB):', self)
        self.fragment_limit_label.setEnabled(False)
        layout.addWidget(self.fragment_limit_label)

        self.fragment_limit_input = QLineEdit(self)
        self.fragment_limit_input.setEnabled(False)
        self.fragment_limit_input.setPlaceholderText('Enter size limit in MB')
        layout.addWidget(self.fragment_limit_input)

        self.setStyleSheet("""
            QLabel {
                color: #FFFFFF;
            }
            QCheckBox {
                color: #FFFFFF;
            }
            QComboBox, QScrollArea {
                background-color: #3E3E3E;
                color: #FFFFFF;
                border: 1px solid #5A5A5A;
                border-radius: 3px;
            }
        """)

    def toggle_fragmentation(self):
        fragmentate = self.fragment_checkbox.isChecked()
        self.fragment_limit_label.setEnabled(fragmentate)
        self.fragment_limit_input.setEnabled(fragmentate)

    def toggle_select_all(self):
        select_all = self.select_all_checkbox.isChecked()
        for i in range(self.scroll_layout.count()):
            checkbox = self.scroll_layout.itemAt(i).widget()
            checkbox.setChecked(select_all)
        self.parent.update_table_preview()

    def update_columns(self, columns):
        self.clear_columns()
        for column in columns:
            checkbox = QCheckBox(column, self)
            checkbox.stateChanged.connect(lambda state, c=column: self.parent.update_table_preview())
            self.scroll_layout.addRow(checkbox)

    def clear_columns(self):
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def get_selected_columns(self):
        return [self.scroll_layout.itemAt(i).widget().text()
                for i in range(self.scroll_layout.count())
                if self.scroll_layout.itemAt(i).widget().isChecked()]

    def get_fragmentation_info(self):
        fragmentate = self.fragment_checkbox.isChecked()
        fragment_limit = None
        if fragmentate:
            fragment_limit = self.fragment_limit_input.text()
            if fragment_limit.isdigit():
                fragment_limit = int(fragment_limit) * 1024 * 1024  # Convert MB to bytes
            else:
                fragment_limit = None
        return fragmentate, fragment_limit

    def close_tab(self):
        self.close_callback(self.file_name)

class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.file_configs = {}
        self.current_file = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('EFC')
        self.setGeometry(100, 100, 1000, 600)

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
            QTabWidget::pane { 
                border-top: 2px solid #C2C7CB;
            }
            QTabWidget::tab-bar {
                left: 5px; 
            }
            QTabBar::tab {
                background: #3E3E3E;
                color: #FFFFFF;
                border: 1px solid #5A5A5A;
                border-bottom-color: #C2C7CB; 
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                min-width: 8ex;
                padding: 2px;
                font-size: 12px; 
            }
            QTabBar::tab:selected, QTabBar::tab:hover {
                background: #5A5A5A;
            }
            QToolButton {
                color: white;
                background-color: red;
                border: none;
                border-radius: 3px;
                padding: 2px;
            }
            QToolButton:hover {
                background-color: darkred;
            }
        """)

        main_layout = QHBoxLayout()
        self.setLayout(main_layout)

        left_layout = QVBoxLayout()
        main_layout.addLayout(left_layout)

        self.input_label = QLabel('Input Folder:', self)
        left_layout.addWidget(self.input_label)

        self.input_line_edit = QLineEdit(self)
        left_layout.addWidget(self.input_line_edit)

        self.input_button = QPushButton('Browse', self)
        self.input_button.clicked.connect(self.browse_input_folder)
        left_layout.addWidget(self.input_button)

        self.output_label = QLabel('Output Folder:', self)
        left_layout.addWidget(self.output_label)

        self.output_line_edit = QLineEdit(self)
        left_layout.addWidget(self.output_line_edit)

        self.output_button = QPushButton('Browse', self)
        self.output_button.clicked.connect(self.browse_output_folder)
        left_layout.addWidget(self.output_button)

        self.tab_widget = QTabWidget(self)
        self.tab_widget.currentChanged.connect(self.update_table_preview)
        left_layout.addWidget(self.tab_widget)

        self.convert_button = QPushButton('Convert', self)
        self.convert_button.clicked.connect(self.convert_files)
        left_layout.addWidget(self.convert_button)

        # Table to display data
        self.table_widget = QTableWidget(self)
        main_layout.addWidget(self.table_widget)

    def browse_input_folder(self):
        options = QFileDialog.Options()
        folder_path = QFileDialog.getExistingDirectory(self, "Select Input Folder", options=options)
        if folder_path:
            self.input_line_edit.setText(folder_path)
            self.update_file_tabs(folder_path)

    def browse_output_folder(self):
        options = QFileDialog.Options()
        folder_path = QFileDialog.getExistingDirectory(self, "Select Output Folder", options=options)
        if folder_path:
            self.output_line_edit.setText(folder_path)

    def update_file_tabs(self, folder_path):
        self.tab_widget.clear()
        self.file_configs.clear()
        self.current_folder = folder_path

        for file_name in os.listdir(folder_path):
            file_path = os.path.normpath(os.path.join(folder_path, file_name))
            if os.path.isfile(file_path) and file_name.lower().endswith(('.xlsx', '.csv', '.json')):
                self.add_file_tab(file_path, file_name)

    def add_closable_tab(self, widget, title):
        tab_index = self.tab_widget.addTab(widget, title)
        tab_button = QToolButton()
        tab_button.setIcon(self.style().standardIcon(QStyle.SP_TitleBarCloseButton))
        tab_button.setIconSize(QSize(12, 12))  # Smaller icon size
        tab_button.setStyleSheet("""
            QToolButton {
                color: white;
                background-color: red;
                border: none;
                border-radius: 3px;
                padding: 2px;
            }
            QToolButton:hover {
                background-color: darkred;
            }
        """)  # Red background and white color
        tab_button.clicked.connect(lambda: self.remove_file_tab(title))
        self.tab_widget.tabBar().setTabButton(tab_index, QTabBar.RightSide, tab_button)
        self.tab_widget.setCurrentIndex(tab_index)

    def add_file_tab(self, file_path, file_name):
        file_path = os.path.normpath(file_path)
        file_config = FileConfig(file_path, file_name, self.remove_file_tab, self)
        self.file_configs[file_path] = file_config

        self.add_closable_tab(file_config, file_name)

        if file_name.lower().endswith('.xlsx'):
            import pandas as pd
            xls = pd.ExcelFile(file_path)
            df = pd.read_excel(file_path, nrows=1)
            file_config.update_columns(df.columns.tolist())
        elif file_name.lower().endswith('.csv'):
            import pandas as pd
            df = pd.read_csv(file_path, nrows=1)
            file_config.update_columns(df.columns.tolist())
        elif file_name.lower().endswith('.json'):
            import json
            import pandas as pd
            data = []
            with open(file_path, 'r', encoding='utf-8') as f:
                for i, line in enumerate(f):
                    if i >= 10:
                        break
                    data.append(json.loads(line.strip()))
            df = pd.json_normalize(data)
            file_config.update_columns(df.columns.tolist())

    def remove_file_tab(self, file_name):
        # Find the full file path using the file_name
        full_file_path = None
        for path in self.file_configs:
            if os.path.basename(path) == file_name:
                full_file_path = path
                break

        if full_file_path is not None:
            index = self.tab_widget.indexOf(self.file_configs[full_file_path])
            if index != -1:
                self.tab_widget.removeTab(index)
                del self.file_configs[full_file_path]

    def convert_files(self):
        output_folder = self.output_line_edit.text()
        if not output_folder:
            QMessageBox.warning(self, "Output Folder Error", "Please select an output folder.")
            return

        for file_path, file_config in self.file_configs.items():
            conversion_type = file_config.type_combo.currentText()
            selected_columns = file_config.get_selected_columns()
            input_file = file_config.file_path

            fragmentate, fragment_limit = file_config.get_fragmentation_info()

            file_name = os.path.basename(file_path)
            output_extension = '.csv' if conversion_type == 'Excel to CSV' or conversion_type == 'JSON to CSV' else '.xlsx'
            output_file = os.path.join(output_folder, os.path.splitext(file_name)[0] + '_converted' + output_extension)

            print(f"Converting {input_file} to {output_file} with columns: {selected_columns} and fragmentation: {fragmentate} with limit: {fragment_limit} bytes")

            try:
                if conversion_type == 'Excel to CSV' and input_file.lower().endswith('.xlsx'):
                    convert_excel(input_file, output_file, selected_columns, fragmentate, fragment_limit)
                elif conversion_type == 'CSV to Excel' and input_file.lower().endswith('.csv'):
                    convert_csv_to_excel(input_file, output_file, selected_columns, fragmentate, fragment_limit)
                elif conversion_type == 'JSON to CSV' and input_file.lower().endswith('.json'):
                    convert_json_to_csv(input_file, output_file, selected_columns, fragmentate, fragment_limit)
                else:
                    QMessageBox.warning(self, "Conversion Type Error", f"Invalid conversion type selected for file {file_name}.")
            except Exception as e:
                QMessageBox.critical(self, "Conversion Error", f"Failed to convert {file_name}: {e}")

        QMessageBox.information(self, "Conversion Complete", "All files have been converted successfully.")

    def update_table_preview(self):
        current_index = self.tab_widget.currentIndex()
        if current_index == -1:
            return
        file_path = os.path.normpath(list(self.file_configs.keys())[current_index])
        selected_columns = [self.file_configs[file_path].scroll_layout.itemAt(i).widget().text()
                            for i in range(self.file_configs[file_path].scroll_layout.count())
                            if self.file_configs[file_path].scroll_layout.itemAt(i).widget().isChecked()]

        if not selected_columns:
            self.table_widget.clear()
            self.table_widget.setRowCount(0)
            self.table_widget.setColumnCount(0)
            return

        if file_path.endswith('.csv'):
            import pandas as pd
            df = pd.read_csv(file_path, usecols=selected_columns, nrows=10)
        elif file_path.endswith('.xlsx'):
            import pandas as pd
            df = pd.read_excel(file_path, usecols=selected_columns, nrows=10)
        elif file_path.endswith('.json'):
            import json
            import pandas as pd
            data = []
            with open(file_path, 'r', encoding='utf-8') as f:
                for i, line in enumerate(f):
                    if i >= 10:
                        break
                    data.append(json.loads(line.strip()))
            df = pd.json_normalize(data)
            df = df[selected_columns]

        self.table_widget.setColumnCount(len(df.columns))
        self.table_widget.setRowCount(len(df.index))
        self.table_widget.setHorizontalHeaderLabels(df.columns)

        for row_index, row_data in df.iterrows():
            for col_index, value in enumerate(row_data):
                self.table_widget.setItem(row_index, col_index, QTableWidgetItem(str(value)))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ConverterApp()
    ex.show()
    sys.exit(app.exec_())