from PyQt6.QtWidgets import QFileDialog
import os

def browse_files(parent, file_types="所有文件 (*);;PDF文件 (*.pdf);;Excel文件 (*.xlsx);;Word文件 (*.docx);;PowerPoint文件 (*.pptx);;LaTeX (*.tex);;Markdown文件 (*.md)"):
    filenames, _ = QFileDialog.getOpenFileNames(parent, "选择文件", "", file_types)
    return filenames

def browse_output_directory(parent):
    directory = QFileDialog.getExistingDirectory(parent, "选择输出目录")
    return directory

def get_default_output_dir(filenames):
    if filenames:
        return os.path.dirname(filenames[0])
    return ""

def ensure_output_directory(output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)