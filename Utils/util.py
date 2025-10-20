import re
from datetime import datetime
from html.parser import HTMLParser


# --- 将 datetime 对象转换为可序列化的字符串 ---
def datetime_to_json(obj):
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError(f"Object of type {type(obj).__name__} is not JSON serializable")

# --- HTML 文本提取器 ---
class HTMLTextExtractor(HTMLParser):
    """用于从 HTML 字符串中提取纯文本的增强型解析器。"""

    def __init__(self):
        super().__init__()
        self.reset()
        self.strict = False
        self.convert_charrefs = True
        self.text = []
        # 新增状态：是否忽略当前数据块 (用于跳过 <style> 和 <script> 内容)
        self.ignore_data = False
        # 定义需要插入换行符的块级标签
        self.block_tags = ('p', 'div', 'br', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'tr', 'li', 'ul', 'ol', 'table')
        self.skip_tags = ('script', 'style', 'head', 'meta', 'title')  # 需要跳过内容的所有标签

    def handle_starttag(self, tag, attrs):
        # 检查是否需要开始忽略数据
        if tag in self.skip_tags:
            self.ignore_data = True

        # 在块级元素开始前插入换行，防止粘连
        if tag in self.block_tags:
            self.text.append('\n')

    def handle_endtag(self, tag):
        # 检查是否需要停止忽略数据
        if tag in self.skip_tags:
            self.ignore_data = False

        # 在块级元素结束后插入换行
        if tag in self.block_tags:
            self.text.append('\n')

    def handle_data(self, data):
        # 只有在不忽略数据的情况下才追加文本
        if not self.ignore_data and data.strip():
            self.text.append(data)

    def get_text(self):
        # 1. 连接所有文本块
        raw_text = ''.join(self.text)

        # 2. 使用正则表达式清理：
        #   - 将多个空格/制表符替换为单个空格
        cleaned_text = re.sub(r'[ \t]+', ' ', raw_text)
        #   - 将多个换行符替换为最多两个 (模拟段落分隔)
        cleaned_text = re.sub(r'\n+', '\n\n', cleaned_text)

        return cleaned_text.strip()


def extract_text_from_html(html_string):
    """
    将 HTML 字符串转换为纯文本。

    Args:
        html_string (str): 包含 HTML 标签的字符串。

    Returns:
        str: 提取出的纯文本内容。
    """
    try:
        extractor = HTMLTextExtractor()
        extractor.feed(html_string)
        # 在解析完成后，手动关闭解析器以处理缓冲区中的内容（确保所有数据都被处理）
        extractor.close()
        return extractor.get_text()
    except Exception as e:
        print(f"警告: HTML 文本提取失败: {e}")
        return html_string  # 失败时返回原始 HTML，交给 AI 处理