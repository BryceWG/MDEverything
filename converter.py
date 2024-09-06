import os
from PyQt6.QtWidgets import QMessageBox
from pdf2markdown import pdf_to_markdown
from excel2markdown import excel_to_markdown
from html2markdown import html_to_markdown, get_webpage_title
from pptx2markdown import pptx_to_markdown
from docx2markdown import docx_to_markdown
import pypandoc
from url_handler import is_url

class Converter:
    def __init__(self, parent, settings_handler):
        self.parent = parent
        self.settings_handler = settings_handler

    def convert_file(self, input_path, output_dir, options):
        """
        根据文件类型调用相应的转换函数。

        :param input_path: 输入文件路径或URL
        :param output_dir: 输出目录
        :param options: 转换选项
        :return: 生成的Markdown文件路径或路径列表
        """
        if is_url(input_path):
            return self.convert_html(input_path, output_dir, options)
        elif os.path.isfile(input_path):
            file_extension = os.path.splitext(input_path)[1].lower()
            converters = {
                ".pdf": self.convert_pdf,
                ".xlsx": self.convert_excel,
                ".pptx": self.convert_pptx,
                ".docx": self.convert_docx_latex,
                ".tex": self.convert_docx_latex
            }
            converter = converters.get(file_extension)
            if converter:
                return converter(input_path, output_dir, options)
            else:
                QMessageBox.warning(self.parent, "警告", f"不支持的文件格式: {input_path}")
        else:
            QMessageBox.warning(self.parent, "警告", f"无效的文件路径或URL: {input_path}")

    def convert_html(self, input_path, output_dir, options):
        try:
            selected_links = options.get('selected_links', [])
            if selected_links:
                return self.convert_multiple_links(selected_links, output_dir, options)
            else:
                return self.convert_single_html(input_path, output_dir, options)
        except Exception as e:
            QMessageBox.critical(self.parent, "错误", f"转换HTML失败: {str(e)}")

    def convert_single_html(self, url, output_dir, options):
        markdown_content = html_to_markdown(
            url,
            use_jina_ai=options.get('use_jina_ai', False),
            jina_api_key=options.get('jina_api_key', ''),
            ignore_links=options.get('ignore_links', False),
            ignore_images=options.get('ignore_images', False),
            body_width=options.get('body_width', None)
        )
        title = get_webpage_title(url)
        safe_title = self.get_safe_filename(title)
        output_file = os.path.join(output_dir, f"{safe_title}.md")
        self.save_markdown(markdown_content, output_file)
        return output_file

    def convert_multiple_links(self, links, output_dir, options):
        all_content = []
        for link in links:
            markdown_content = html_to_markdown(
                link,
                use_jina_ai=options.get('use_jina_ai', False),
                jina_api_key=options.get('jina_api_key', ''),
                ignore_links=options.get('ignore_links', False),
                ignore_images=options.get('ignore_images', False),
                body_width=options.get('body_width', None)
            )
            title = get_webpage_title(link)
            all_content.append(f"# {title}\n\n{markdown_content}\n\n---\n\n")
        
        combined_content = "".join(all_content)
        safe_title = self.get_safe_filename("combined_webpages")
        output_file = os.path.join(output_dir, f"{safe_title}.md")
        self.save_markdown(combined_content, output_file)
        return output_file

    def convert_pdf(self, input_path, output_dir, options):
        try:
            self.settings_handler.save_pdf_settings(options['app_id'], options['secret_code'])
            markdown_content = pdf_to_markdown(input_path, **options)
            output_file = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(input_path))[0]}.md")
            self.save_markdown(markdown_content, output_file)
            return output_file
        except Exception as e:
            QMessageBox.critical(self.parent, "PDF转换错误", f"PDF转换失败: {str(e)}\n\n请检查pdf2markdown.py文件是否正确配置。")

    def convert_excel(self, input_path, output_dir, options):
        selected_sheets = options['selected_sheets']
        if not selected_sheets:
            QMessageBox.warning(self.parent, "警告", f"未选择 {os.path.basename(input_path)} 的工作表")
            return

        converted_files = []
        for sheet_name in selected_sheets:
            output_file = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(input_path))[0]}-{sheet_name}.md")
            try:
                excel_to_markdown(input_path, output_file, sheet_name, options['has_header'])
                converted_files.append(output_file)
            except Exception as e:
                QMessageBox.critical(self.parent, "错误", f"转换工作表 '{sheet_name}' 失败: {str(e)}")
        return converted_files

    def convert_pptx(self, input_path, output_dir, options):
        try:
            return pptx_to_markdown(input_path, output_dir, **options)
        except Exception as e:
            QMessageBox.critical(self.parent, "错误", f"PPTX转换失败: {str(e)}")

    def convert_docx_latex(self, input_path, output_dir, options):
        try:
            output_file = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(input_path))[0]}.md")
            if input_path.lower().endswith('.docx'):
                docx_to_markdown(input_path, output_file)
            else:  # .tex 文件
                pypandoc.convert_file(input_path, 'md', outputfile=output_file, format='latex')
            return output_file
        except Exception as e:
            self.show_conversion_error(e, input_path)

    def save_markdown(self, content, filename):
        try:
            with open(filename, "w", encoding="utf-8") as f:
                f.write(content)
            QMessageBox.information(self.parent, "成功", f"Markdown文件已保存为: {filename}")
        except Exception as e:
            QMessageBox.critical(self.parent, "错误", f"保存Markdown文件失败: {str(e)}")

    @staticmethod
    def get_safe_filename(filename):
        return "".join([c for c in filename if c.isalnum() or c in (' ', '-', '_')]).rstrip()

    def show_conversion_error(self, error, input_path):
        error_message = f"转换失败: {str(error)}\n"
        if input_path.lower().endswith('.tex'):
            error_message += "\n请确保已正确安装pypandoc和pandoc。\n"
            error_message += "可以尝试运行以下命令安装:\n"
            error_message += "pip install pypandoc\n"
            error_message += "并从 https://pandoc.org/installing.html 下载安装pandoc"
        QMessageBox.critical(self.parent, "错误", error_message)