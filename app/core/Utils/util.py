import re
from datetime import datetime, timezone
from email.utils import getaddresses
from html.parser import HTMLParser

MIN_AWARE_DATETIME = datetime.min.replace(tzinfo=timezone.utc)

# --- 将 datetime 对象转换为可序列化的字符串 ---
def datetime_to_json(obj):
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError(f"Object of type {type(obj).__name__} is not JSON serializable")


def get_sortable_time(email_data):
    """
    一个健壮的排序键，
    可处理 sent_time 是 str、datetime 或 None 的情况。
    """

    sent_time = email_data.get("sent_time")

    if isinstance(sent_time, datetime):
        # 如果是 datetime，确保它有TzInfo
        if sent_time.tzinfo is None:
            # 假设 naive (无时区) 的时间是 UTC
            return sent_time.replace(tzinfo=timezone.utc)
        return sent_time  # 它已是 datetime 对象, 直接返回

    if isinstance(sent_time, str):
        try:
            # 尝试将其从 ISO 字符串转为 datetime 对象
            return datetime.fromisoformat(sent_time)
        except (ValueError, TypeError):
            # 如果字符串格式错误，返回一个最小值
            return MIN_AWARE_DATETIME

    # 如果 sent_time 是 None 或其他类型, 返回最小值
    return MIN_AWARE_DATETIME


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


def get_address_list_from_header(headers):
    """
        从邮件头部字段 (To, Cc) 中解析并提取所有邮箱地址。

        Args:
            header_value (str): 原始邮件头部字段值。

        Returns:
            list: 邮箱地址字符串列表。
        """
    if not headers:
        return []

    # 使用 email.utils.getaddresses 来处理复杂的地址字符串，包括编码和多个地址
    addresses = getaddresses([headers])

    # addresses 是一个 (display_name, email_address) 元组的列表
    return [email_addr for display_name, email_addr in addresses if email_addr]


# --- 辅助函数 (用于归档, 避免重复代码) ---
def archive_email_to_memory(email, all_memory, processed_ids_this_run):
    """
    将一封已格式化的邮件归档到 all_memory 字典中。
    处理 received 和 sent 两种情况, 并处理ID去重。
    """
    email_id = email.get("id")

    # 检查此邮件是否 *在本次运行中* 已被添加
    # (这主要用于防止一封“已发送”邮件被多次计数)
    if email_id and email_id in processed_ids_this_run:
        return 0  # 已添加, 返回 0

    other_party_addresses = []
    email_type = email.get("type")

    if email_type == "received":
        # (已在 Step 2 中格式化)
        if email.get("sender"):
            other_party_addresses.append(email["sender"])

    elif email_type == "sent":
        receivers = email.get("receiver")
        if isinstance(receivers, list):
            other_party_addresses.extend(receivers)

    if not other_party_addresses:
        print(f"警告：邮件 {email_id} 无法确定归档地址，跳过。")
        return 0

    # 将邮件添加到所有相关的对方地址下
    for address in set(other_party_addresses):
        if not address: continue

        if address not in all_memory:
            all_memory[address] = []

        # (我们假设此邮件的 ID 已经通过了 existing_ids 的检查)
        all_memory[address].append(email)

    if email_id:
        processed_ids_this_run.add(email_id)  # 标记此ID在本次运行中已处理

    return 1  # 成功添加 1 封