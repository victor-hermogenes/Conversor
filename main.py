import sys
import os
import threading
import json
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, 
    QPushButton, QFileDialog, QComboBox, QMessageBox, QCheckBox, 
    QScrollArea, QFormLayout, QTableWidget, QTableWidgetItem, QHBoxLayout, QTabWidget, QToolButton, QStyle, QTabBar, QProgressDialog, QDialog, QDialogButtonBox
)
from PyQt5.QtCore import Qt, QSize, pyqtSignal, QObject
from PyQt5.QtGui import QIcon
from functions import convert_excel, convert_json_to_csv, convert_csv_to_excel, fragment_file
import pandas as pd

class WorkerSignals(QObject):
    progress = pyqtSignal(int, int)
    complete = pyqtSignal()

class WorkerThread(threading.Thread):
    def __init__(self, file_configs, output_folder, fragment_size_mb, signals):
        super().__init__()
        self.file_configs = file_configs
        self.output_folder = output_folder
        self.fragment_size_mb = fragment_size_mb
        self.signals = signals

    def run(self):
        total_files = len(self.file_configs)
        for i, (file_path, file_config) in enumerate(self.file_configs.items()):
            conversion_type = file_config.type_combo.currentText()
            selected_columns = file_config.get_selected_columns()
            input_file = file_config.file_path
            delimiter = file_config.delimiter_combo.currentText()
            string_delimiter = file_config.string_delimiter_line_edit.text()

            file_name = os.path.basename(file_path)
            output_extension = '.csv' if conversion_type == 'Excel to CSV' or conversion_type == 'JSON to CSV' else '.xlsx'
            output_file = os.path.join(self.output_folder, os.path.splitext(file_name)[0] + '_converted' + output_extension)

            try:
                if conversion_type == 'Excel to CSV' and input_file.lower().endswith('.xlsx'):
                    convert_excel(input_file, output_file, selected_columns)
                elif conversion_type == 'CSV to Excel' and input_file.lower().endswith('.csv'):
                    convert_csv_to_excel(input_file, output_file, selected_columns, delimiter, string_delimiter)
                elif conversion_type == 'JSON to CSV' and input_file.lower().endswith('.json'):
                    convert_json_to_csv(input_file, output_file, selected_columns)
                else:
                    continue

                if self.fragment_size_mb:
                    fragment_file(output_file, self.fragment_size_mb)

            except Exception as e:
                print(f"Failed to convert {file_name}: {e}")

            self.signals.progress.emit(i + 1, total_files)

        self.signals.complete.emit()

class FileConfig(QWidget):
    def __init__(self, file_path, file_name, close_callback, parent):
        super().__init__(parent)
        self.file_path = os.path.normpath(file_path)
        self.file_name = file_name
        self.close_callback = close_callback
        self.parent = parent
        self.all_columns = []  # Store all columns here for filtering
        self.column_checkboxes = {}  # Store checkboxes for columns
        self.original_order = []  # Store the original order of column names
        self.initUI()

    def initUI(self):
        self.setWindowIcon(QIcon('conversor.ico'))
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

        self.search_bar = QLineEdit(self)
        self.search_bar.setPlaceholderText('Search columns...')
        self.search_bar.textChanged.connect(self.filter_columns)
        layout.addWidget(self.search_bar)

        self.select_all_checkbox = QCheckBox('Select All', self)
        self.select_all_checkbox.stateChanged.connect(self.toggle_select_all)
        layout.addWidget(self.select_all_checkbox)

        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget(self.scroll_area)
        self.scroll_layout = QFormLayout(self.scroll_content)
        self.scroll_area.setWidget(self.scroll_content)
        layout.addWidget(self.scroll_area)

        self.delimiter_label = QLabel('Delimiter (for CSV to Excel):', self)
        layout.addWidget(self.delimiter_label)

        self.delimiter_combo = QComboBox(self)
        self.delimiter_combo.addItems([",", ";", "\t", "|"])
        layout.addWidget(self.delimiter_combo)

        self.string_delimiter_label = QLabel('String Delimiter (for CSV to Excel):', self)
        layout.addWidget(self.string_delimiter_label)

        self.string_delimiter_line_edit = QLineEdit(self)
        self.string_delimiter_line_edit.setText('"')
        layout.addWidget(self.string_delimiter_line_edit)

        self.copy_selection_button = QPushButton('Copy Selection To', self)
        self.copy_selection_button.clicked.connect(self.copy_selection_to)
        layout.addWidget(self.copy_selection_button)

        self.setStyleSheet("""
            QLabel {
                color: #FFFFFF;
            }
            QCheckBox {
                color: #FFFFFF;
            }
            QLineEdit, QComboBox, QScrollArea, QPushButton {
                background-color: #3E3E3E;
                color: #FFFFFF;
                border: 1px solid #5A5A5A;
                border-radius: 3px;
            }
        """)

    def filter_columns(self):
        search_text = self.search_bar.text().lower()
        if search_text == '':
            for column in self.original_order:
                self.column_checkboxes[column].setParent(None)
                self.scroll_layout.addRow(self.column_checkboxes[column])
                self.column_checkboxes[column].show()
        else:
            for column, checkbox in self.column_checkboxes.items():
                checkbox.setParent(None)
                if search_text in column.lower():
                    self.scroll_layout.addRow(checkbox)
                    checkbox.show()
                else:
                    checkbox.hide()

    def toggle_select_all(self):
        select_all = self.select_all_checkbox.isChecked()
        for checkbox in self.column_checkboxes.values():
            if checkbox.isVisible():
                checkbox.setChecked(select_all)
        self.parent.update_table_preview()

    def update_columns(self, columns):
        self.clear_columns()
        self.all_columns = columns
        self.original_order = columns[:]
        for column in columns:
            checkbox = QCheckBox(column, self)
            checkbox.stateChanged.connect(lambda state, c=column: self.parent.update_table_preview())
            self.column_checkboxes[column] = checkbox
            self.scroll_layout.addRow(checkbox)

    def clear_columns(self):
        for i in reversed(range(self.scroll_layout.count())):
            widget = self.scroll_layout.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)

    def get_selected_columns(self):
        return [column for column, checkbox in self.column_checkboxes.items() if checkbox.isChecked()]

    def close_tab(self):
        self.close_callback(self.file_name)

    def copy_selection_to(self):
        selected_columns = self.get_selected_columns()
        dialog = CopySelectionDialog(self.parent, selected_columns)
        if dialog.exec_() == QDialog.Accepted:
            target_sheets = dialog.get_selected_sheets()
            for sheet_name in target_sheets:
                if sheet_name in self.parent.file_configs:
                    self.parent.file_configs[sheet_name].set_columns(selected_columns)
        self.parent.update_table_preview()

    def set_columns(self, selected_columns):
        for column, checkbox in self.column_checkboxes.items():
            checkbox.setChecked(column in selected_columns)
        self.parent.update_table_preview()

class CopySelectionDialog(QDialog):
    def __init__(self, parent, selected_columns):
        super().__init__(parent)
        self.setWindowTitle('Copy Selection To')
        self.setWindowIcon(QIcon('conversor.ico'))
        self.selected_columns = selected_columns
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        self.setLayout(layout)

        self.sheets_label = QLabel('Select sheets to copy the selection to:', self)
        layout.addWidget(self.sheets_label)

        self.select_all_checkbox = QCheckBox('Select All', self)
        self.select_all_checkbox.stateChanged.connect(self.toggle_select_all)
        layout.addWidget(self.select_all_checkbox)

        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget(self.scroll_area)
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_area.setWidget(self.scroll_content)
        layout.addWidget(self.scroll_area)

        self.sheet_checkboxes = {}
        for sheet_name, file_config in self.parent().file_configs.items():
            checkbox = QCheckBox(sheet_name, self)
            self.sheet_checkboxes[sheet_name] = checkbox
            self.scroll_layout.addWidget(checkbox)

        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def toggle_select_all(self):
        select_all = self.select_all_checkbox.isChecked()
        for checkbox in self.sheet_checkboxes.values():
            checkbox.setChecked(select_all)

    def get_selected_sheets(self):
        return [sheet for sheet, checkbox in self.sheet_checkboxes.items() if checkbox.isChecked()]

class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.file_configs = {}
        self.current_file = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('EFC')
        self.setWindowIcon(QIcon('conversor.ico'))
        self.resize_to_screen()
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

        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_area.setWidget(scroll_widget)
        scroll_layout = QVBoxLayout(scroll_widget)
        main_layout = QHBoxLayout()
        self.setLayout(main_layout)
        main_layout.addWidget(scroll_area)

        left_layout = QVBoxLayout()
        scroll_layout.addLayout(left_layout)

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

        self.fragment_checkbox = QCheckBox('Enable Fragmentation', self)
        self.fragment_checkbox.stateChanged.connect(self.toggle_fragmentation)
        left_layout.addWidget(self.fragment_checkbox)

        self.fragment_size_label = QLabel('Fragment Size (MB):', self)
        left_layout.addWidget(self.fragment_size_label)

        self.fragment_size_line_edit = QLineEdit(self)
        self.fragment_size_line_edit.setPlaceholderText('Enter fragment size in MB')
        self.fragment_size_line_edit.setEnabled(False)
        left_layout.addWidget(self.fragment_size_line_edit)

        self.example_checkbox = QCheckBox('Example sheet', self)
        self.example_checkbox.stateChanged.connect(self.toggle_example_sheet)
        left_layout.addWidget(self.example_checkbox)

        self.example_file_edit = QLineEdit(self)
        self.example_file_edit.setPlaceholderText('Select example sheet...')
        self.example_file_edit.setEnabled(False)
        left_layout.addWidget(self.example_file_edit)

        self.example_file_button = QPushButton('Browse and Read', self)
        self.example_file_button.setEnabled(False)
        self.example_file_button.clicked.connect(self.browse_example_sheet)
        left_layout.addWidget(self.example_file_button)

        self.tab_widget = QTabWidget(self)
        self.tab_widget.currentChanged.connect(self.update_table_preview)
        left_layout.addWidget(self.tab_widget)

        self.convert_button = QPushButton('Convert', self)
        self.convert_button.clicked.connect(self.convert_files)
        left_layout.addWidget(self.convert_button)

        self.table_widget = QTableWidget(self)
        main_layout.addWidget(self.table_widget)

    def resize_to_screen(self):
        screen_geometry = QApplication.desktop().screenGeometry()
        self.setGeometry(screen_geometry)
        self.showMaximized()

    def toggle_fragmentation(self):
        self.fragment_size_line_edit.setEnabled(self.fragment_checkbox.isChecked())

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

    def toggle_example_sheet(self):
        self.example_file_edit.setEnabled(self.example_checkbox.isChecked())
        self.example_file_button.setEnabled(self.example_checkbox.isChecked())

    def browse_example_sheet(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Example Sheet", "", "Excel Files (*.xlsx *.xls);;CSV Files (*.csv)", options=options)
        if file_path:
            self.example_file_edit.setText(file_path)
            self.process_example_sheet(file_path)

    def process_example_sheet(self, file_path):
        try:
            if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                df = pd.read_excel(file_path)
            elif file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                raise ValueError("Unsupported file type. Please upload a CSV or Excel file.")

            columns = [col.strip() for col in df.columns]  # Trimming column headers
            for file_config in self.file_configs.values():
                for column, checkbox in file_config.column_checkboxes.items():
                    if column.strip() in columns:
                        checkbox.setChecked(True)

            self.update_table_preview()
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to process example sheet: {e}')

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
        tab_button.setIconSize(QSize(12, 12))
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
        """)
        tab_button.clicked.connect(lambda: self.remove_file_tab(widget))
        self.tab_widget.tabBar().setTabButton(tab_index, QTabBar.RightSide, tab_button)
        self.tab_widget.setCurrentIndex(tab_index)

    def add_file_tab(self, file_path, file_name):
        file_path = os.path.normpath(file_path)
        file_config = FileConfig(file_path, file_name, self.remove_file_tab, self)
        self.file_configs[file_path] = file_config

        self.add_closable_tab(file_config, file_name)

        if file_name.lower().endswith('.xlsx'):
            df = pd.read_excel(file_path, nrows=1)
            file_config.update_columns(df.columns.tolist())
        elif file_name.lower().endswith('.csv'):
            df = pd.read_csv(file_path, nrows=1)
            file_config.update_columns(df.columns.tolist())
        elif file_name.lower().endswith('.json'):
            data = []
            with open(file_path, 'r', encoding='utf-8') as f:
                for i, line in enumerate(f):
                    if i >= 10:
                        break
                    data.append(json.loads(line.strip()))
            df = pd.json_normalize(data)
            file_config.update_columns(df.columns.tolist())

    def remove_file_tab(self, file_config_widget):
        for file_path, file_config in self.file_configs.items():
            if file_config == file_config_widget:
                index = self.tab_widget.indexOf(file_config_widget)
                if index != -1:
                    self.tab_widget.removeTab(index)
                    del self.file_configs[file_path]
                break

    def convert_files(self):
        output_folder = self.output_line_edit.text()
        fragment_size = self.fragment_size_line_edit.text()
        
        # Validation for fragment size input
        if self.fragment_checkbox.isChecked():
            try:
                fragment_size_mb = float(fragment_size)
            except ValueError:
                QMessageBox.warning(self, "Fragment Size Error", "Please enter a valid number for the fragment size.")
                return
        else:
            fragment_size_mb = None

        if not output_folder:
            QMessageBox.warning(self, "Output Folder Error", "Please select an output folder.")
            return

        self.progress_dialog = QProgressDialog("Converting files...", "Cancel", 0, len(self.file_configs), self)
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.setMinimumDuration(0)
        self.progress_dialog.setValue(0)

        self.signals = WorkerSignals()
        self.signals.progress.connect(self.update_progress)
        self.signals.complete.connect(self.conversion_complete)

        self.worker_thread = WorkerThread(self.file_configs, output_folder, fragment_size_mb, self.signals)
        self.worker_thread.start()

    def update_progress(self, value, total):
        self.progress_dialog.setValue(value)
        if self.progress_dialog.wasCanceled():
            self.worker_thread.join(0)

    def conversion_complete(self):
        self.progress_dialog.setValue(len(self.file_configs))
        QMessageBox.information(self, "Conversion Complete", "All files have been converted successfully.")

    def update_table_preview(self):
        current_index = self.tab_widget.currentIndex()
        if current_index == -1:
            return
        file_path = list(self.file_configs.keys())[current_index]
        selected_columns = [self.file_configs[file_path].scroll_layout.itemAt(i).widget().text()
                            for i in range(self.file_configs[file_path].scroll_layout.count())
                            if self.file_configs[file_path].scroll_layout.itemAt(i).widget().isChecked()]

        if not selected_columns:
            self.table_widget.clear()
            self.table_widget.setRowCount(0)
            self.table_widget.setColumnCount(0)
            return

        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path, usecols=selected_columns, nrows=10)
        elif file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, usecols=selected_columns, nrows=10)
        elif file_path.endswith('.json'):
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
    app.setWindowIcon(QIcon('conversor.ico'))
    ex = ConverterApp()
    ex.show()
    sys.exit(app.exec_())