import html2text
import requests
import re
from urllib.parse import urlparse, urljoin
from bs4 import BeautifulSoup
from PyQt6.QtWidgets import QMessageBox, QProgressDialog, QListWidgetItem
from PyQt6.QtCore import Qt, QObject, pyqtSignal, QRunnable

class WorkerSignals(QObject):
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

class LinkExtractorWorker(QRunnable):
    def __init__(self, url):
        super().__init__()
        self.url = url
        self.signals = WorkerSignals()

    def run(self):
        try:
            links = self.extract_links(self.url)
            self.signals.finished.emit(links)
        except Exception as e:
            self.signals.error.emit(str(e))

    def extract_links(self, url):
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            base_url = '{uri.scheme}://{uri.netloc}'.format(uri=urlparse(url))
            links = []
            for a in soup.find_all('a', href=True):
                href = a['href']
                full_url = urljoin(base_url, href)
                if urlparse(full_url).netloc == urlparse(base_url).netloc:
                    title = a.text.strip() or full_url
                    links.append((title, full_url))
            return links
        except Exception as e:
            print(f"提取链接时出错: {str(e)}")
            return []

class HTMLHandler:
    def __init__(self, parent, threadpool):
        self.parent = parent
        self.threadpool = threadpool

    def load_webpage_links(self, url, links_list):
        """加载网页链接并更新链接列表"""
        if not self.is_url(url):
            QMessageBox.warning(self.parent, "警告", "请输入有效的URL")
            return

        progress = QProgressDialog("正在加载链接...", "取消", 0, 0, self.parent)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.show()

        worker = LinkExtractorWorker(url)
        worker.signals.finished.connect(lambda links: self.update_links_list(links, links_list, progress))
        worker.signals.error.connect(lambda error: self.show_error(error, progress))
        self.threadpool.start(worker)

    def update_links_list(self, links, links_list, progress):
        """更新链接列表UI"""
        progress.close()
        links_list.clear()
        for title, link in links:
            item = QListWidgetItem(f"{title} ({link})")
            item.setData(Qt.ItemDataRole.UserRole, link)
            links_list.addItem(item)
        
        if not links:
            QMessageBox.information(self.parent, "信息", "未找到任何链接")

    def show_error(self, error, progress):
        """显示错误消息"""
        progress.close()
        QMessageBox.critical(self.parent, "错误", f"加载链接失败: {error}")

    @staticmethod
    def is_url(text):
        """检查文本是否为有效URL"""
        url_pattern = re.compile(
            r'^(?:http|ftp)s?://'
            r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+(?:[A-Z]{2,6}\.?|[A-Z0-9-]{2,}\.?)|'
            r'localhost|'
            r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})'
            r'(?::\d+)?'
            r'(?:/?|[/?]\S+)$', re.IGNORECASE)
        return url_pattern.match(text) is not None

def html_to_markdown(url, use_jina_ai=False, jina_api_key=None, ignore_links=False, ignore_images=False, body_width=None):
    """
    将HTML转换为Markdown格式。

    :param url: 要转换的网页URL
    :param use_jina_ai: 是否使用Jina AI进行转换
    :param jina_api_key: Jina AI的API密钥
    :param ignore_links: 是否忽略链接
    :param ignore_images: 是否忽略图片
    :param body_width: 正文宽度
    :return: 转换后的Markdown内容
    """
    if use_jina_ai:
        return jina_html_to_markdown(url, jina_api_key)
    else:
        return standard_html_to_markdown(url, ignore_links, ignore_images, body_width)

def standard_html_to_markdown(url, ignore_links, ignore_images, body_width):
    """使用标准html2text库进行转换"""
    try:
        response = requests.get(url)
        response.raise_for_status()
        html = response.text

        h = html2text.HTML2Text()
        h.ignore_links = ignore_links
        h.ignore_images = ignore_images
        if body_width is not None:
            h.body_width = body_width
        
        return h.handle(html)
    except requests.RequestException as e:
        raise ConnectionError(f"获取网页内容失败: {str(e)}")

def jina_html_to_markdown(url, api_key):
    """使用Jina AI进行转换"""
    if not api_key:
        raise ValueError("Jina API密钥不能为空")

    headers = {
        'Authorization': f'Bearer {api_key}'
    }
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.text
    except requests.RequestException as e:
        raise ConnectionError(f"Jina AI请求失败: {str(e)}")

def get_webpage_title(url):
    """获取网页标题"""
    try:
        response = requests.get(url)
        response.raise_for_status()
        html = response.text
        start = html.find('<title>') + 7
        end = html.find('</title>')
        if start > 6 and end != -1:
            return html[start:end].strip()
        return urlparse(url).netloc
    except requests.RequestException:
        return urlparse(url).netloc

# 使用示例
if __name__ == "__main__":
    test_url = "https://example.com"
    try:
        markdown_content = html_to_markdown(test_url)
        print("转换成功!")
        print("Markdown内容:")
        print(markdown_content)
    except Exception as e:
        print(f"转换失败: {str(e)}")