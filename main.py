import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QMessageBox, 
                             QComboBox, QCheckBox, QGroupBox, QFormLayout, QTabWidget, QSpinBox, QListWidget, QTreeWidget, QTreeWidgetItem, QListWidgetItem, QProgressDialog)
from PyQt6.QtCore import Qt, QThreadPool
from PyQt6.QtGui import QPalette, QColor, QIcon
import os
from url_handler import is_url
from excel2markdown import update_excel_sheets_tree
from markdown_merger import merge_markdown_files
from file_handler import browse_files, browse_output_directory, get_default_output_dir, ensure_output_directory
from settings_handler import SettingsHandler
from converter import Converter
from html2markdown import HTMLHandler

class MarkdownConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.init_connections()
        
        # 设置窗口图标
        self.set_app_icon()

    def set_app_icon(self):
        icon_path = self.get_icon_path()
        if icon_path:
            try:
                icon = QIcon(icon_path)
                if icon.isNull():
                    print(f"警告：图标文件加载失败: {icon_path}")
                else:
                    self.setWindowIcon(icon)
                    print(f"成功设置图标: {icon_path}")
            except Exception as e:
                print(f"设置图标时发生错误: {str(e)}")
        else:
            print("警告：无法找到图标文件")

    def get_icon_path(self):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        print(f"当前目录: {current_dir}")
        icon_names = ["MDEverything.ico", "MDEverything.png"]
        
        for icon_name in icon_names:
            icon_path = os.path.join(current_dir, icon_name)
            print(f"尝试加载图标: {icon_path}")
            if os.path.exists(icon_path):
                print(f"找到图标: {icon_path}")
                return icon_path
        
        print("未找到任何图标文件")
        return None

    def init_ui(self):
        self.setWindowTitle("多格式转换Markdown工具")
        self.setGeometry(100, 100, 600, 500)

        self.settings_handler = SettingsHandler("YourCompany", "MarkdownConverter")

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        main_layout.addWidget(self.create_file_selection_group())
        main_layout.addWidget(self.create_output_directory_group())
        main_layout.addWidget(self.create_options_tab())
        main_layout.addWidget(self.create_convert_button())

        self.set_app_style()

        self.threadpool = QThreadPool()
        self.converter = Converter(self, self.settings_handler)
        self.html_handler = HTMLHandler(self, self.threadpool)

    def init_connections(self):
        self.file_entry.editingFinished.connect(self.check_input_type)
        self.browse_button.clicked.connect(self.browse_file)
        self.output_browse_button.clicked.connect(self.browse_output_directory)
        self.convert_button.clicked.connect(self.convert)
        self.use_jina_ai.stateChanged.connect(self.toggle_html_options)
        self.load_links_button.clicked.connect(self.load_webpage_links)

    def create_file_selection_group(self):
        group = QGroupBox("文件选择")
        layout = QVBoxLayout()
        
        file_input_layout = QHBoxLayout()
        self.file_entry = QLineEdit()
        self.file_entry.setPlaceholderText("输入文件路径或URL")
        self.browse_button = QPushButton("浏览文件")
        file_input_layout.addWidget(self.file_entry)
        file_input_layout.addWidget(self.browse_button)
        
        layout.addLayout(file_input_layout)
        group.setLayout(layout)
        return group

    def create_output_directory_group(self):
        group = QGroupBox("输出目录")
        layout = QHBoxLayout()
        self.output_entry = QLineEdit()
        self.output_browse_button = QPushButton("浏览")
        layout.addWidget(self.output_entry)
        layout.addWidget(self.output_browse_button)
        group.setLayout(layout)
        return group

    def create_options_tab(self):
        self.options_tab = QTabWidget()
        self.options_tab.addTab(self.create_pdf_options(), "PDF选项")
        self.options_tab.addTab(self.create_html_options(), "HTML选项")
        self.options_tab.addTab(self.create_docx_latex_options(), "Word/LaTeX选项")
        self.options_tab.addTab(self.create_pptx_options(), "PPT选项")
        self.options_tab.addTab(self.create_excel_options(), "Table选项")
        self.options_tab.addTab(self.create_markdown_merge_options(), "Markdown合并")
        return self.options_tab

    def create_convert_button(self):
        self.convert_button = QPushButton("转换")
        self.set_button_style(self.convert_button)
        return self.convert_button

    def set_app_style(self):
        QApplication.setStyle("Fusion")  # 将样式应用到整个应用程序
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Highlight, QColor(58, 119, 52))
        palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)
        self.setPalette(palette)

    def set_button_style(self, button):
        button.setMinimumHeight(30)
        button.setStyleSheet("""
            QPushButton {
                padding: 5px 10px;
                font-size: 14px;
            }
        """)

    def create_pdf_options(self):
        pdf_group = QGroupBox("PDF转换选项")
        pdf_layout = QFormLayout()

        self.app_id_entry = QLineEdit()
        self.app_id_entry.setEchoMode(QLineEdit.EchoMode.Password)
        self.app_id_entry.setText(self.settings_handler.load_setting("app_id", ""))
        pdf_layout.addRow("x-ti-app-id:", self.app_id_entry)

        self.secret_code_entry = QLineEdit()
        self.secret_code_entry.setEchoMode(QLineEdit.EchoMode.Password)
        self.secret_code_entry.setText(self.settings_handler.load_setting("secret_code", ""))
        pdf_layout.addRow("x-ti-secret-code:", self.secret_code_entry)

        self.dpi_combo = QComboBox()
        self.dpi_combo.addItems(["72", "144", "216"])
        self.dpi_combo.setCurrentText("216")
        pdf_layout.addRow("DPI:", self.dpi_combo)

        self.parse_mode_combo = QComboBox()
        self.parse_mode_combo.addItems(["auto", "scan"])
        pdf_layout.addRow("Parse Mode:", self.parse_mode_combo)

        self.apply_document_tree = QCheckBox("生成标题")
        self.apply_document_tree.setChecked(True)
        pdf_layout.addRow(self.apply_document_tree)

        self.table_flavor = QComboBox()
        self.table_flavor.addItems(["md", "html"])
        pdf_layout.addRow("表格格式:", self.table_flavor)

        self.get_image = QComboBox()
        self.get_image.addItems(["none", "page", "objects", "both"])
        pdf_layout.addRow("获取图片:", self.get_image)

        self.page_start = QSpinBox()
        self.page_start.setMinimum(1)
        pdf_layout.addRow("起始页:", self.page_start)

        self.page_count = QSpinBox()
        self.page_count.setMinimum(1)
        self.page_count.setMaximum(1000)
        self.page_count.setValue(1000)
        pdf_layout.addRow("页数:", self.page_count)

        pdf_group.setLayout(pdf_layout)
        return pdf_group

    def create_excel_options(self):
        excel_group = QGroupBox("Table转换选项")
        excel_layout = QVBoxLayout()

        self.excel_sheets_tree = QTreeWidget()
        self.excel_sheets_tree.setHeaderLabel("Excel文件和工作表")
        self.excel_sheets_tree.setSelectionMode(QTreeWidget.SelectionMode.MultiSelection)
        self.excel_sheets_tree.setStyleSheet("""
            QTreeWidget::item:selected {
                background-color: #3A7734;
                color: white;
            }
        """)
        excel_layout.addWidget(QLabel("选择工作表:"))
        excel_layout.addWidget(self.excel_sheets_tree)

        self.excel_header = QCheckBox("第一行为表头")
        self.excel_header.setChecked(True)
        excel_layout.addWidget(self.excel_header)

        excel_group.setLayout(excel_layout)
        return excel_group

    def create_html_options(self):
        html_group = QGroupBox("HTML转换选项")
        html_layout = QVBoxLayout()

        self.jina_api_key = QLineEdit()
        self.jina_api_key.setEchoMode(QLineEdit.EchoMode.Password)
        self.jina_api_key.setPlaceholderText("输入Jina API密钥")
        html_layout.addWidget(QLabel("Jina API密钥:"))
        html_layout.addWidget(self.jina_api_key)

        self.use_jina_ai = QCheckBox("使用Jina AI")
        self.use_jina_ai.stateChanged.connect(self.toggle_html_options)
        html_layout.addWidget(self.use_jina_ai)

        options_layout = QFormLayout()
        self.html_ignore_links = QCheckBox("忽略链接")
        options_layout.addRow(self.html_ignore_links)

        self.html_ignore_images = QCheckBox("忽略图片")
        options_layout.addRow(self.html_ignore_images)

        self.html_body_width = QSpinBox()
        self.html_body_width.setMinimum(0)
        self.html_body_width.setMaximum(1000)
        self.html_body_width.setValue(0)
        options_layout.addRow("正文宽度:", self.html_body_width)

        html_layout.addLayout(options_layout)

        self.load_links_button = QPushButton("加载网页链接")
        self.load_links_button.clicked.connect(self.load_webpage_links)
        html_layout.addWidget(self.load_links_button)

        self.links_list = QListWidget()
        self.links_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        html_layout.addWidget(self.links_list)

        html_group.setLayout(html_layout)
        return html_group

    def toggle_html_options(self, state):
        enabled = not bool(state)
        self.html_ignore_links.setEnabled(enabled)
        self.html_ignore_images.setEnabled(enabled)
        self.html_body_width.setEnabled(enabled)
        self.links_list.setEnabled(enabled)
        self.load_links_button.setEnabled(enabled)

    def load_webpage_links(self):
        url = self.file_entry.text().strip()
        self.html_handler.load_webpage_links(url, self.links_list)

    def create_pptx_options(self):
        pptx_group = QGroupBox("PPT转换选项")
        pptx_layout = QFormLayout()

        self.pptx_image_width = QSpinBox()
        self.pptx_image_width.setMinimum(0)
        self.pptx_image_width.setMaximum(2000)
        self.pptx_image_width.setValue(800)
        pptx_layout.addRow("图片最大宽度 (px):", self.pptx_image_width)

        self.pptx_disable_image = QCheckBox("禁用图片提取")
        pptx_layout.addRow(self.pptx_disable_image)

        self.pptx_disable_escaping = QCheckBox("不转义特殊字符")
        pptx_layout.addRow(self.pptx_disable_escaping)

        self.pptx_disable_notes = QCheckBox("不添加演讲者注释")
        pptx_layout.addRow(self.pptx_disable_notes)

        self.pptx_disable_color = QCheckBox("禁用颜色标签")
        pptx_layout.addRow(self.pptx_disable_color)

        self.pptx_enable_slides = QCheckBox("启用幻灯片分隔符")
        pptx_layout.addRow(self.pptx_enable_slides)

        self.pptx_min_block_size = QSpinBox()
        self.pptx_min_block_size.setMinimum(0)
        self.pptx_min_block_size.setMaximum(1000)
        self.pptx_min_block_size.setValue(0)
        pptx_layout.addRow("最小文本块大小:", self.pptx_min_block_size)

        self.pptx_output_format = QComboBox()
        self.pptx_output_format.addItems(["markdown", "wiki", "mdk", "qmd"])
        pptx_layout.addRow("输出格式:", self.pptx_output_format)

        pptx_group.setLayout(pptx_layout)
        return pptx_group

    def create_docx_latex_options(self):
        docx_latex_group = QGroupBox("Word/LaTeX转换选项")
        docx_latex_layout = QFormLayout()

        self.docx_latex_format = QComboBox()
        self.docx_latex_format.addItems(["docx", "latex"])
        docx_latex_layout.addRow("输入格式:", self.docx_latex_format)

        docx_latex_group.setLayout(docx_latex_layout)
        return docx_latex_group

    def create_markdown_merge_options(self):
        markdown_merge_group = QGroupBox("Markdown合并选项")
        markdown_merge_layout = QVBoxLayout()

        self.markdown_files_list = QListWidget()
        self.markdown_files_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        markdown_merge_layout.addWidget(QLabel("选择要合并的Markdown文件:"))
        markdown_merge_layout.addWidget(self.markdown_files_list)

        merge_button = QPushButton("合并选中的Markdown文件")
        merge_button.clicked.connect(self.merge_markdown_files)
        markdown_merge_layout.addWidget(merge_button)

        markdown_merge_group.setLayout(markdown_merge_layout)
        return markdown_merge_group

    def update_markdown_files_list(self, filenames):
        self.markdown_files_list.clear()
        for filename in filenames:
            if filename.lower().endswith('.md'):
                item = QListWidgetItem(os.path.basename(filename))
                item.setData(Qt.ItemDataRole.UserRole, filename)
                self.markdown_files_list.addItem(item)

    def merge_markdown_files(self):
        selected_items = self.markdown_files_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请选择要合并的Markdown文件")
            return

        output_dir = self.output_entry.text()
        if not output_dir:
            QMessageBox.warning(self, "警告", "请选择输出目录")
            return

        selected_files = [item.data(Qt.ItemDataRole.UserRole) for item in selected_items]
        
        try:
            output_file = merge_markdown_files(selected_files, output_dir)
            QMessageBox.information(self, "成功", f"合并的Markdown文件已保存为: {output_file}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"合并Markdown文件失败: {str(e)}")

    def browse_file(self):
        filenames = browse_files(self)
        if filenames:
            self.file_entry.setText(";".join(filenames))
            default_output_dir = get_default_output_dir(filenames)
            self.output_entry.setText(default_output_dir)
            
            self.update_markdown_files_list(filenames)
            
            if len(filenames) == 1:
                self.update_options_tab(filenames[0])
            
            update_excel_sheets_tree(self.excel_sheets_tree, filenames)
        
        self.file_entry.setMinimumHeight(30)
        self.output_entry.setMinimumHeight(30)

    def check_input_type(self):
        input_text = self.file_entry.text().strip()
        if is_url(input_text):
            self.options_tab.setCurrentIndex(1)
        elif ";" in input_text:
            pass
        else:
            self.update_options_tab(input_text)

    def update_options_tab(self, filename):
        if is_url(filename):
            self.options_tab.setCurrentIndex(1)
        else:
            file_extension = os.path.splitext(filename)[1].lower()
            if file_extension == ".pdf":
                self.options_tab.setCurrentIndex(0)
            elif file_extension == ".xlsx":
                self.options_tab.setCurrentIndex(4)
            elif file_extension in [".html", ".htm"]:
                self.options_tab.setCurrentIndex(1)
            elif file_extension == ".pptx":
                self.options_tab.setCurrentIndex(3)
            elif file_extension in [".docx", ".tex"]:
                self.options_tab.setCurrentIndex(2)

    def browse_output_directory(self):
        directory = browse_output_directory(self)
        if directory:
            self.output_entry.setText(directory)

    def convert(self):
        input_paths = self.file_entry.text().split(";")
        output_dir = self.output_entry.text()
        
        if not output_dir:
            QMessageBox.warning(self, "警告", "请选择输出目录")
            return
        
        try:
            ensure_output_directory(output_dir)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"创建输出目录失败: {str(e)}")
            return

        converted_files = []
        for input_path in input_paths:
            input_path = input_path.strip()
            options = self.get_conversion_options()
            result = self.converter.convert_file(input_path, output_dir, options)
            if result:
                if isinstance(result, list):
                    converted_files.extend(result)
                else:
                    converted_files.append(result)

        if converted_files:
            if len(converted_files) > 1:
                QMessageBox.information(self, "完成", f"所有文件转换完成，共转换 {len(converted_files)} 个文件。")
            else:
                QMessageBox.information(self, "完成", f"文件已转换: {converted_files[0]}")
        else:
            QMessageBox.warning(self, "警告", "没有文件被成功转换")

    def get_conversion_options(self):
        options = {
            'use_jina_ai': self.use_jina_ai.isChecked(),
            'jina_api_key': self.jina_api_key.text(),
            'ignore_links': self.html_ignore_links.isChecked(),
            'ignore_images': self.html_ignore_images.isChecked(),
            'body_width': self.html_body_width.value(),
            'app_id': self.app_id_entry.text(),
            'secret_code': self.secret_code_entry.text(),
            'dpi': int(self.dpi_combo.currentText()),
            'apply_document_tree': int(self.apply_document_tree.isChecked()),
            'table_flavor': self.table_flavor.currentText(),
            'get_image': self.get_image.currentText(),
            'page_start': self.page_start.value(),
            'page_count': self.page_count.value(),
            'parse_mode': self.parse_mode_combo.currentText(),
            'selected_sheets': self.get_selected_excel_sheets(),
            'has_header': self.excel_header.isChecked(),
            'image_width': self.pptx_image_width.value(),
            'disable_image': self.pptx_disable_image.isChecked(),
            'disable_escaping': self.pptx_disable_escaping.isChecked(),
            'disable_notes': self.pptx_disable_notes.isChecked(),
            'disable_color': self.pptx_disable_color.isChecked(),
            'enable_slides': self.pptx_enable_slides.isChecked(),
            'min_block_size': self.pptx_min_block_size.value(),
            'output_format': self.pptx_output_format.currentText(),
        }
        
        selected_links = [item.data(Qt.ItemDataRole.UserRole) for item in self.links_list.selectedItems()]
        if selected_links:
            options['selected_links'] = selected_links
        
        return options

    def get_selected_excel_sheets(self):
        selected_items = self.excel_sheets_tree.selectedItems()
        return [item.data(0, Qt.ItemDataRole.UserRole) for item in selected_items if item.parent()]

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    window = MarkdownConverterApp()
    app_icon = window.windowIcon()
    if app_icon.isNull():
        print("警告：应用程序图标为空")
    else:
        app.setWindowIcon(app_icon)
        print("成功设置应用程序图标")
    
    window.show()
    
    app.aboutToQuit.connect(window.threadpool.waitForDone)
    
    sys.exit(app.exec())