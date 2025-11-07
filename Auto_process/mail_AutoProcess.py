import imaplib
import email
import json
import os
import threading
import pytz
import AI_Handler
import time as time_module

from datetime import datetime, time
from email.utils import parseaddr
from email.header import decode_header
from email.utils import parsedate_to_datetime
from google import genai
from Utils.util import datetime_to_json, extract_text_from_html, get_address_list_from_header, archive_email_to_memory, \
    get_sortable_time

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
MAIL_CONFIG_FILE = os.path.join(CURRENT_DIR, "../Configs/Setup/mail_config.json")
AI_CONFIG_FILE = os.path.join(CURRENT_DIR,"../Configs/Setup/AI_config.json")
IN_RAWDATA_OUTPUT_PATH = os.path.join(CURRENT_DIR, "../Info/inbox_data.json")
SENT_RAWDATA_OUTPUT_PATH = os.path.join(CURRENT_DIR, "../Info/sentbox_data.json")
SCORE_LIST_PATH = os.path.join(CURRENT_DIR, "../Info/email_sender_score_list.json")
VALID_MAIL_OUTPUT_PATH = os.path.join(CURRENT_DIR, "../Info/valid_emails.json")
INVALID_MAIL_OUTPUT_PATH = os.path.join(CURRENT_DIR, "../Info/invalid_emails.json")
SENT_MAIL_OUTPUT_PATH = os.path.join(CURRENT_DIR, "../Info/sent_emails.json")
CONVERSATION_MEMORY_PATH = os.path.join(CURRENT_DIR, "../Info/conversation_memory.json")

# 读取邮箱配置
try:
    with open(MAIL_CONFIG_FILE, 'r', encoding='utf-8') as f:
        EMAIL_CONFIG = json.load(f)['EMAIL_CONFIG']
except FileNotFoundError:
    print(f"错误：找不到配置文件 {MAIL_CONFIG_FILE}，请检查路径。")
    exit()
except json.JSONDecodeError:
    print(f"错误：配置文件 {MAIL_CONFIG_FILE} 格式不正确。")
    exit()

# --- 邮箱配置 ---
IMAP_SERVER = EMAIL_CONFIG['IMAP_SERVER']
IMAP_PORT = EMAIL_CONFIG['IMAP_PORT']
EMAIL_ADDRESS = EMAIL_CONFIG['EMAIL_ADDRESS']
APP_PASSWORD = EMAIL_CONFIG['APP_PASSWORD']
VALID_SCORE = EMAIL_CONFIG['THRESHOLD']['VALID_SCORE']
NO_REPLY_PATTERN = EMAIL_CONFIG['NO_REPLY_PATTERNS']

# --- 时区设置 ---
TIMEZONE_STR = EMAIL_CONFIG['TIME_AREA'] + "/" + EMAIL_CONFIG['TIME_NATION']
TIMEZONE = pytz.timezone(TIMEZONE_STR)


# 读取AI配置
try:
    with open(AI_CONFIG_FILE, 'r', encoding='utf-8') as f:
        AI_CONFIG = json.load(f)["GEMINI_API"]
except FileNotFoundError:
    print(f"错误：找不到配置文件 {MAIL_CONFIG_FILE}，请检查路径。")
    exit()
except json.JSONDecodeError:
    print(f"错误：配置文件 {MAIL_CONFIG_FILE} 格式不正确。")
    exit()


# --- AI配置 ---

API_KEY = AI_CONFIG['API_KEY']
MODEL_NAME = AI_CONFIG['MODEL_NAME']


# --- 连接到IMAP服务器并登录 ---
def connect_and_login_email():
    mclient = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mclient.login(EMAIL_ADDRESS, APP_PASSWORD)
    mclient.select('inbox')
    return mclient


# --- 连接GEMINIAPI ---
def connect_gemini():
    return genai.Client(api_key=API_KEY)


# --- 读取未读邮件,结构化并保存为原始数据 ---
def fetch_unseen_emails(mclient, json_file_path=IN_RAWDATA_OUTPUT_PATH):
    # 搜索所有未读 (UNSEEN) 邮件
    status, email_ids = mclient.search(None, 'UNSEEN')
    email_id_list = email_ids[0].split()

    emails = []

    for email_id in email_id_list:
        # 获取邮件的完整数据 (RFC822 格式)
        # 注意：这里 status, msg_data 的获取需要在循环内部
        # mail.fetch() 的第一个返回值是状态，第二个是数据
        status, msg_data = mclient.fetch(email_id, '(RFC822)')

        msg = email.message_from_bytes(msg_data[0][1])

        date_header = msg['Date']

        try:
            sent_time = parsedate_to_datetime(date_header)
        except Exception:
            sent_time = None  # 遇到格式错误时设置为 None

        subject_tuple = decode_header(msg['Subject'])[0]
        subject = subject_tuple[0].decode(subject_tuple[1] or 'utf-8') if isinstance(subject_tuple[0], bytes) else \
        subject_tuple[0]

        body = ""
        html_body = ""

        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                cdispo = str(part.get('Content-Disposition'))
                payload = part.get_payload(decode=True)

                if ctype == 'text/plain' and 'attachment' not in cdispo and payload:
                    # 找到纯文本，优先使用，并跳出循环
                    body = payload.decode('utf-8', errors='ignore').strip()
                    if body:
                        break

                elif ctype == 'text/html' and 'attachment' not in cdispo and payload:
                    # 存储 HTML 内容作为后备
                    html_body = payload.decode('utf-8', errors='ignore').strip()

        else:
            # 非 multipart 消息
            payload = msg.get_payload(decode=True)
            if msg.get_content_type() == 'text/plain' and payload:
                body = payload.decode('utf-8', errors='ignore').strip()
            elif msg.get_content_type() == 'text/html' and payload:
                html_body = payload.decode('utf-8', errors='ignore').strip()

        if len(body) == 0:
            body = extract_text_from_html(html_body)

        # ----------------------------------------------------
        # 转换为指定时区时间
        if sent_time:
            sent_time_jst = sent_time.astimezone(TIMEZONE)
        else:
            sent_time_jst = None

        sender = msg['From']
        to_header = msg['To']
        cc_header = msg['Cc']

        to = get_address_list_from_header(to_header)
        cc = get_address_list_from_header(cc_header)

        # 解析sender
        sender_root = ''
        sender_name = ''

        if sender:
            display_name, email_addr = parseaddr(sender)
            sender_name, sender_root = email_addr.split('@')


        # 结构化返回信息
        emails.append({
            'type': 'received',
            'id': email_id.decode(),
            'sender_root': sender_root,
            'sender_name': sender_name,
            'receiver': to,
            'cc': cc,
            'subject': subject,
            'sent_time': sent_time_jst,
            'body': body
        })

        # 处理完后标记为已读
        mclient.store(email_id, '+FLAGS', '\\Seen')

    # ------------------- JSON 写入部分 -------------------
    if emails:

        # 1. 尝试读取现有数据
        all_emails = []
        try:
            if os.path.exists(json_file_path) and os.path.getsize(json_file_path) > 0:
                with open(json_file_path, 'r', encoding='utf-8') as f:
                    all_emails = json.load(f)
        except Exception as e:
            # 如果文件存在但读取失败 (例如 JSON 格式错误)，打印警告并继续
            print(f"WARNING: 原始数据文件读取失败 ({e})，将以新数据覆盖。")
            all_emails = []

            # 2. 合并新数据,并按发送时间进行排序
        all_emails.extend(emails)
        all_emails.sort(key=lambda email: email['sent_time']
                 if isinstance(email['sent_time'], datetime)
                 else datetime.fromisoformat(email['sent_time']))

        # 3. 写入完整合并后的数据
        with open(json_file_path, 'w', encoding='utf-8') as f:
            json.dump(all_emails, f, indent=4, ensure_ascii=False, default=datetime_to_json)
        print(f"成功提取 {len(emails)} 封邮件，并写入到 {json_file_path}")
    else:
        print("没有发现新的未读邮件。")

    return emails


# --- 读取发送邮件,结构化并保存为原始数据 ---
def fetch_sent_emails(mclient, json_file_path=SENT_RAWDATA_OUTPUT_PATH):
    # 1. 选择已发送文件夹。如果 'Sent' 失败，可以尝试 'Sent Items' 或其他特定名称
    try:
        status, messages = mclient.select('Sent')
    except mclient.error:
        print("警告：无法选择 'Sent' 文件夹。请检查您的邮箱服务商是否使用了其他名称（如 'Sent Items'）。")
        return []

    # 2. 搜索所有邮件 (已发送邮件通常被视为已读，不能用 UNSEEN)
    status, email_ids = mclient.search(None, 'ALL')
    email_id_list = email_ids[0].split()

    emails = []

    for email_id in email_id_list:
        # 获取邮件的完整数据 (RFC822 格式)
        status, msg_data = mclient.fetch(email_id, '(RFC822)')

        msg = email.message_from_bytes(msg_data[0][1])

        date_header = msg['Date']

        try:
            sent_time = parsedate_to_datetime(date_header)
        except Exception:
            sent_time = None

        # 获取 Subject
        subject_tuple = decode_header(msg['Subject'])[0]
        subject = subject_tuple[0].decode(subject_tuple[1] or 'utf-8') if isinstance(subject_tuple[0], bytes) else \
            subject_tuple[0]

        body = ""
        html_body = ""

        # 邮件正文提取逻辑 (与 fetch_unseen_emails 相同)
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                cdispo = str(part.get('Content-Disposition'))
                payload = part.get_payload(decode=True)

                if ctype == 'text/plain' and 'attachment' not in cdispo and payload:
                    body = payload.decode('utf-8', errors='ignore').strip()
                    if body:
                        break

                elif ctype == 'text/html' and 'attachment' not in cdispo and payload:
                    html_body = payload.decode('utf-8', errors='ignore').strip()

        else:
            payload = msg.get_payload(decode=True)
            if msg.get_content_type() == 'text/plain' and payload:
                body = payload.decode('utf-8', errors='ignore').strip()
            elif msg.get_content_type() == 'text/html' and payload:
                html_body = payload.decode('utf-8', errors='ignore').strip()

        if not body and html_body:
            body = extract_text_from_html(html_body)

        # ----------------------------------------------------
        # 转换为指定时区时间
        if sent_time:
            sent_time_jst = sent_time.astimezone(TIMEZONE)
        else:
            sent_time_jst = None

        # 获取 Sender
        sender = msg['From']

        # 获取 To 和 Cc
        to_header = msg.get('To')
        cc_header = msg.get('Cc')

        to = get_address_list_from_header(to_header)
        cc = get_address_list_from_header(cc_header)

        display_name, email_addr = parseaddr(sender)


        # 结构化返回信息
        emails.append({
            'type': 'sent',
            'id': email_id.decode(),
            'sender': email_addr,
            'receiver': to,
            'cc': cc,
            'subject': subject,
            'sent_time': sent_time_jst,
            'body': body
        })

    # ------------------- JSON 写入部分 (新增去重逻辑) -------------------
    new_unique_emails = []

    if emails:

        # 尝试读取现有数据
        all_emails = []
        existing_ids = set()  # 存储现有邮件的ID
        try:
            if os.path.exists(json_file_path) and os.path.getsize(json_file_path) > 0:
                with open(json_file_path, 'r', encoding='utf-8') as f:
                    all_emails = json.load(f)

                # 从已读取的数据中收集所有 ID
                existing_ids = {email.get('id') for email in all_emails if email.get('id')}
        except Exception as e:
            print(f"WARNING: 原始数据文件读取失败 ({e})，将以新数据覆盖。")
            all_emails = []


        if not len(all_emails) > 0:
            new_unique_emails = emails
        else:
            new_unique_emails = [email for email in emails if email.get('id') not in existing_ids]

        all_emails.extend(new_unique_emails)  # 只追加不重复的邮件

        # 排序逻辑
        all_emails.sort(key=lambda email: email['sent_time']
        if isinstance(email['sent_time'], datetime)
        else datetime.fromisoformat(email['sent_time']))

        # 写入完整合并后的数据
        with open(json_file_path, 'w', encoding='utf-8') as f:
            json.dump(all_emails, f, indent=4, ensure_ascii=False, default=datetime_to_json)
        print(f"成功提取 {len(new_unique_emails)} 封新增的已发送邮件，并写入到 {json_file_path}")
    else:
        print("没有发现新的已发送邮件。")

    return new_unique_emails


# --- 对邮件分类并存储，随后根据该发件地址对分数列表进行维护 ---
def email_classification(ai_client, in_emails, sent_emails, invalid_output_path=INVALID_MAIL_OUTPUT_PATH, valid_output_path=VALID_MAIL_OUTPUT_PATH, sent_output_path=SENT_MAIL_OUTPUT_PATH):
    valid_emails = []
    invalid_emails = []
    uncertain_emails = []

    if not len(sent_emails) > 0:
        sent_emails = []


    if not (in_emails or sent_emails):
        print("无邮件需要分类")
        return valid_emails, sent_emails

    # 收到邮件处理
    if len(in_emails) > 0:
        try:
            with open(SCORE_LIST_PATH, 'r', encoding='utf-8') as f:
                mail_score_file = json.load(f)
                score_list = mail_score_file["SENDER_INFO_LIST"]
        except FileNotFoundError:
            print(f"错误：找不到配置文件 {SCORE_LIST_PATH}，请检查路径。")
            score_list = {}
        except json.JSONDecodeError:
            print(f"错误：配置文件 {SCORE_LIST_PATH} 格式不正确。")
            score_list = {}

        os.makedirs(os.path.dirname(invalid_output_path), exist_ok=True)

        # 遍历读取的邮件，根据地址名单命中情况与具体评分进行分类
        for email in in_emails:
            sender_root = email['sender_root']
            sender_name = email['sender_name']

            sender_name_list = score_list.get(sender_root)
            if sender_name_list is None:
                uncertain_emails.append(email)
                continue

            score = sender_name_list.get(sender_name)
            if score is None:
                uncertain_emails.append(email)

            # 若存在评分，则为其数据结构统一添加
            elif score >= VALID_SCORE:
                email["score"] = score
                valid_emails.append(email)

            else :
                email["score"] = score
                invalid_emails.append(email)


        # 完成分类后分别进行对应处理
        # 未识别邮件交由AI根据摘要和内容进行评分后分为有效和无效邮件中
        if len(uncertain_emails) > 0:
            print(f"SUCCESS: {len(uncertain_emails)} 封邮件被初步筛选为待定,等待后续识别归档")
            # 交由AI读取其内容并为其进行评分,返回邮件字典-评分的元组的列表
            result_list = AI_Handler.get_score_for_uncertain_emails(ai_client, uncertain_emails, MODEL_NAME)

            count = 0
            # 根据结果字典维护邮件评分文件
            if result_list:
                for email in result_list:
                    sender_root = email['sender_root']
                    sender_name = email['sender_name']

                    if sender_root not in score_list:
                        score_list[sender_root] = {}

                    if sender_name in score_list[sender_root]:
                        score_list[sender_root][sender_name] = int(round((score_list[sender_root][sender_name] + email["score"]) / 2))
                    else:
                        score_list[sender_root][sender_name] = email["score"]

                    count += 1

                # 将结果重新写入评分表文件
                with open(SCORE_LIST_PATH, 'w', encoding='utf-8') as f:
                    score_output = {"SENDER_INFO_LIST":score_list}
                    json.dump(score_output, f, indent=4, ensure_ascii=False, default=datetime_to_json)
                    print(
                        f"SUCCESS: {count} 条记录被维护到 {SCORE_LIST_PATH}中")

                # 进行邮件有效性的区分
                for email in result_list:
                    if email["score"] >= VALID_SCORE:
                        valid_emails.append(email)
                    else:
                        invalid_emails.append(email)
                print(f"SUCCESS: {len(uncertain_emails)} 封邮件重分类成功")


        # 无效邮件直接进行存储
        if len(invalid_emails) > 0:
            # 1. 读取现有无效邮件数据
            all_invalid_emails = []
            try:
                if os.path.exists(invalid_output_path) and os.path.getsize(invalid_output_path) > 0:
                    with open(invalid_output_path, 'r', encoding='utf-8') as f:
                        all_invalid_emails = json.load(f)
            except Exception as e:
                print(f"WARNING: 读取现有无效邮件文件失败 ({e})，将以新数据开始写入。")
                all_invalid_emails = []

            # 2. 合并新数据
            all_invalid_emails.extend(invalid_emails)
            all_invalid_emails.sort(key=lambda email: email['sent_time']
                     if isinstance(email['sent_time'], datetime)
                     else datetime.fromisoformat(email['sent_time']))

            # 3. 覆盖写入完整列表
            with open(invalid_output_path, 'w', encoding='utf-8') as f:
                json.dump(all_invalid_emails, f, indent=4, ensure_ascii=False, default=datetime_to_json)
                print(f"SUCCESS: {len(invalid_emails)} 封邮件被标记为无效邮件，写入 {invalid_output_path}")


        # 有效邮件交由AI获取总结,完善其数据结构,后续用于对话记忆等功能
        if len(valid_emails) > 0:
            # 对于没有总结的，交由AI进行总结
            need_summarize_list = []
            for email in valid_emails:
                if not email.get("summary",None):
                    need_summarize_list.append(email)

            if len(need_summarize_list) > 0:
                # 由于是对数据源的引用，所以更改数据源后所有引用无需手动更新
                AI_Handler.get_summary_for_emails(ai_client, need_summarize_list, MODEL_NAME)

            # 存储有效邮件
            # 1. 读取现有有效邮件数据
            all_valid_emails = []
            try:
                if os.path.exists(valid_output_path) and os.path.getsize(valid_output_path) > 0:
                    with open(valid_output_path, 'r', encoding='utf-8') as f:
                        all_valid_emails = json.load(f)
            except Exception as e:
                print(f"WARNING: 读取现有有效邮件文件失败 ({e})，将以新数据开始写入。")
                all_valid_emails = []

            # 2. 合并新数据
            all_valid_emails.extend(valid_emails)
            all_valid_emails.sort(key=lambda email: email['sent_time']
                     if isinstance(email['sent_time'], datetime)
                     else datetime.fromisoformat(email['sent_time']))

            # 3. 覆盖写入完整列表
            with open(valid_output_path, 'w', encoding='utf-8') as f:
                json.dump(all_valid_emails, f, indent=4, ensure_ascii=False, default=datetime_to_json)
                print(f"SUCCESS: {len(valid_emails)} 封邮件被标记为有效邮件，将在处理后写入 {valid_output_path},并用于记忆构成")


    # 发送邮件处理
    if len(sent_emails) > 0:
        sent_bol = True
        print(f"SUCCESS: {len(sent_emails)} 封已发送邮件将交由AI总结内容")
        # 交由AI进行总结
        AI_Handler.get_summary_for_emails(ai_client, sent_emails, MODEL_NAME)

        # 存储发送邮件
        all_sent_emails = []
        try:
            if os.path.exists(sent_output_path) and os.path.getsize(sent_output_path) > 0:
                with open(sent_output_path, 'r', encoding='utf-8') as f:
                    all_sent_emails = json.load(f)
        except Exception as e:
            print(f"WARNING: 读取现有有效邮件文件失败 ({e})，将以新数据开始写入。")
            all_sent_emails = []

        # 合并新数据
        all_sent_emails.extend(sent_emails)
        all_sent_emails.sort(key=lambda email: email['sent_time']
        if isinstance(email['sent_time'], datetime)
        else datetime.fromisoformat(email['sent_time']))

        # 覆盖写入完整列表
        with open(sent_output_path, 'w', encoding='utf-8') as f:
            json.dump(all_sent_emails, f, indent=4, ensure_ascii=False, default=datetime_to_json)
            print(f"SUCCESS: {len(sent_emails)} 封发送邮件，将写入 {sent_output_path},并用于记忆构成")


    return valid_emails, sent_emails


# --- 根据历史邮件构建对话历史 ---
def init_conversation_history(ai_client, all_valid_emails_path=VALID_MAIL_OUTPUT_PATH,
                              all_sent_emails_path=SENT_MAIL_OUTPUT_PATH, memory_file_path=CONVERSATION_MEMORY_PATH):
    print("\n//////////////////对话历史初始化...//////////////////")

    # --- 1. (约束检查) ---
    if os.path.exists(memory_file_path):
        try:
            if os.path.getsize(memory_file_path) > 10:
                with open(memory_file_path, 'r', encoding='utf-8') as f:
                    content = json.load(f)
                if content:
                    print(f"信息：对话历史文件 {memory_file_path} 已存在且不为空。终止初始化。")
                    print("//////////////////对话历史初始化终止。//////////////////\n")
                    return
        except Exception as e:
            print(f"警告：对话历史文件 {memory_file_path} 存在但无法解析({e})。终止初始化。")
            print("//////////////////对话历史初始化终止。//////////////////\n")
            return

    print(f"信息：对话历史文件为空，开始从历史邮件构建...")

    # --- 2. 加载所有历史邮件 ---
    try:
        with open(all_valid_emails_path, 'r', encoding='utf-8') as f:
            all_valid_emails = json.load(f)
        with open(all_sent_emails_path, 'r', encoding='utf-8') as f:
            all_sent_emails = json.load(f)
    except FileNotFoundError as e:
        print(f"FATAL: 找不到历史邮件文件 {e.filename}。初始化失败。")
        print("//////////////////对话历史初始化失败。//////////////////\n")
        return
    except Exception as e:
        print(f"FATAL: 无法读取历史邮件文件: {e}")
        print("//////////////////对话历史初始化失败。//////////////////\n")
        return

    print(f"信息：已加载 {len(all_valid_emails)} 封收信和 {len(all_sent_emails)} 封发信。")

    # --- 3. (发信优先) 格式化 sent_emails ---
    formatted_sent_emails = []
    addresses_from_sent_mail = set()

    if len(all_sent_emails) > 0:
        for email in all_sent_emails:
            receivers = email.get("receiver")
            if not isinstance(receivers, list) or not receivers:
                continue
            formatted_sent_emails.append(email)
            addresses_from_sent_mail.update(receivers)

    print(f"信息：(发信优先) 从 {len(formatted_sent_emails)} 封发信中提取了 {len(addresses_from_sent_mail)} 个对话地址。")

    # --- 4. (格式化收信) ---
    formatted_valid_emails = []
    if len(all_valid_emails) > 0:
        for email in all_valid_emails:
            if any(pattern in email.get("sender_name", "").lower() for pattern in NO_REPLY_PATTERN):
                continue
            sender_addr = email.get("sender_name", "unknown") + "@" + email.get("sender_root", "unknown.com")
            if not sender_addr or sender_addr == "unknown@unknown.com":
                continue
            email["sender"] = sender_addr
            email.pop("sender_name", None)
            email.pop("sender_root", None)
            email.pop("score", None)
            formatted_valid_emails.append(email)
        print(f"信息：历史收信格式化完成，{len(formatted_valid_emails)} 封邮件有效。")

    # --- 5. (分流) ---
    known_addresses = addresses_from_sent_mail
    emails_to_add_fast = []  # 快速通道
    emails_to_filter_slow = []  # 慢速通道

    for email in formatted_valid_emails:
        if email["sender"] in known_addresses:
            emails_to_add_fast.append(email)
        else:
            emails_to_filter_slow.append(email)

    emails_to_add_fast.extend(formatted_sent_emails)
    print(f"信息：邮件分流完成。快速通道: {len(emails_to_add_fast)} 封，慢速(AI)通道: {len(emails_to_filter_slow)} 封。")

    # --- 6. (慢速通道) 运行 AI 清洗 ---
    if emails_to_filter_slow:
        print(f"信息：正在提交 {len(emails_to_filter_slow)} 封新邮件到 AI 进行内容清洗...")
        print("(这可能需要很长时间，取决于邮件数量...)")

        filtered_new_emails = AI_Handler.get_conversation_constitutes_for_emails(
            ai_client, emails_to_filter_slow
        )
        print(f"信息：AI 清洗完成，{len(filtered_new_emails)} 封邮件被确认为新对话。")
        emails_to_add_fast.extend(filtered_new_emails)
    else:
        print("信息：没有需要 AI 清洗的邮件。")

    # --- 7. (归档) ---
    all_emails_to_archive = emails_to_add_fast

    if not all_emails_to_archive:
        print("警告：没有可用于归档的邮件。初始化终止。")
        return

    print(f"信息：开始归档 {len(all_emails_to_archive)} 封最终邮件...")

    # (注意：all_memory 此时是旧的数据结构: {"address": [email_list]})
    all_memory = {}
    new_email_added_count = 0
    processed_ids_this_run = set()

    for email in all_emails_to_archive:
        new_email_added_count += archive_email_to_memory(
            email, all_memory, processed_ids_this_run
        )

    print(f"信息：归档完成。总共添加了 {new_email_added_count} 封邮件到 {len(all_memory)} 条对话中。")

    # --- 8. (排序、生成总结/口吻并保存) ---
    print("信息：正在对所有对话进行时间排序...")
    for email_list in all_memory.values():
        if not isinstance(email_list, list): continue
        try:
            email_list.sort(key=get_sortable_time)
        except Exception as e:
            print(f"警告：对话 {e} 排序失败。")

    print("信息：调用AI为所有对话生成总体总结...")
    memory_with_summaries = AI_Handler.get_history_summary_for_conversation(ai_client, all_memory)

    # (修改点 2: 将上一步的结果 传入 第二个 AI 函数)
    print("信息：调用AI为所有对话生成口吻分析...")
    final_memory_structure = AI_Handler.get_style_profile_for_conversation(ai_client, memory_with_summaries)

    print(f"信息：AI 分析与数据结构合并完成。")

    # (修改点 3: 保存)
    try:
        with open(memory_file_path, 'w', encoding='utf-8') as f:
            # (保存最终的、包含总结和口吻的完整结构)
            json.dump(final_memory_structure, f, ensure_ascii=False, indent=2, default=datetime_to_json)
        print(f"信息：对话历史已成功初始化并保存到 {memory_file_path}")
    except Exception as e:
        print(f"错误：保存对话历史文件失败 ({e})")

    print("//////////////////对话历史初始化完成。//////////////////\n")


# --- 根据获取的有效邮件维护对话历史 ---
def maintain_conversation_history(ai_client, valid_emails, sent_emails, memory_file_path=CONVERSATION_MEMORY_PATH):
    # 步骤：
    # 1. 读取对话历史 (修正)
    # 2. (发信优先) 格式化 sent_emails
    # 3. (发信优先) 结合 known_addresses
    # 4. (格式化收信) 格式化 valid_emails
    # 5. (分流)
    # 6. (慢速通道) 运行 AI 清洗
    # 7. (归档) (修正) *内联* 新的归档逻辑
    # 8. (排序) (修正) 排序新结构
    # 9. (总结与口吻) (实现) *仅* 为 'updated_addresses' 调用 AI 管道
    # 10. (保存)

    if not (valid_emails or sent_emails):
        print("无可用于对话历史维护的邮件")
        return
    print("\n//////////////////开始对话历史维护...//////////////////")

    # --- 1. 读取对话历史记录 (修正) ---
    all_memory = {}
    try:
        if os.path.exists(memory_file_path) and os.path.getsize(memory_file_path) > 0:
            with open(memory_file_path, 'r', encoding='utf-8') as f:
                all_memory = json.load(f)
                if not isinstance(all_memory, dict):
                    all_memory = {}
    except Exception as e:
        print(f"警告：读取现有对话历史文件失败 ({e})，将从零开始构建。")
        all_memory = {}

    existing_ids = set()
    for convo_data in all_memory.values():
        if isinstance(convo_data, dict):
            email_list = convo_data.get("emails", [])
            for email in email_list:
                if email and email.get("id"):
                    existing_ids.add(email.get("id"))

    existing_addresses_from_memory = set(all_memory.keys())
    print(f"信息：已加载 {len(existing_ids)} 个邮件ID 和 {len(existing_addresses_from_memory)} 条已有对话。")

    # --- 2. (发信优先) 格式化 sent_emails ---
    formatted_sent_emails = []
    addresses_from_sent_mail = set()

    if len(sent_emails) > 0:
        for email in sent_emails:
            if email.get("id") and email.get("id") in existing_ids:
                continue
            receivers = email.get("receiver")
            if not isinstance(receivers, list) or not receivers:
                continue
            formatted_sent_emails.append(email)
            addresses_from_sent_mail.update(receivers)

    # --- 3. (发信优先) 结合 "已知地址" ---
    known_addresses = existing_addresses_from_memory.union(addresses_from_sent_mail)

    if len(addresses_from_sent_mail) > 0:
        print(f"信息：(发信优先) 本次发送的邮件中新增了 {len(addresses_from_sent_mail)} 个已知地址。")
        print(f"信息：已知对话伙伴总数更新为 {len(known_addresses)}。")

    # --- 4. (格式化收信) 处理 valid_emails ---
    formatted_valid_emails = []
    if len(valid_emails) > 0:
        for email in valid_emails:
            if email.get("id") and email.get("id") in existing_ids:
                continue
            if any(pattern in email.get("sender_name", "").lower() for pattern in NO_REPLY_PATTERN):
                continue

            sender_addr = email.get("sender_name", "unknown") + "@" + email.get("sender_root", "unknown.com")
            if not sender_addr or sender_addr == "unknown@unknown.com":
                continue

            email["sender"] = sender_addr
            email.pop("sender_name", None)
            email.pop("sender_root", None)
            email.pop("score", None)

            formatted_valid_emails.append(email)
        print(f"信息：收信格式化完成，{len(formatted_valid_emails)} 封邮件有效。")

    # --- 5. (分流) ---
    emails_to_add_fast = []  # 快速通道
    emails_to_filter_slow = []  # 慢速通道

    for email in formatted_valid_emails:
        if email["sender"] in known_addresses:
            emails_to_add_fast.append(email)
        else:
            emails_to_filter_slow.append(email)

    emails_to_add_fast.extend(formatted_sent_emails)
    print(f"信息：邮件分流完成。快速通道: {len(emails_to_add_fast)} 封，慢速(AI)通道: {len(emails_to_filter_slow)} 封。")

    # --- 6. (慢速通道) 运行 AI 清洗 ---
    if emails_to_filter_slow:
        print(f"信息：正在提交 {len(emails_to_filter_slow)} 封新邮件到 AI 进行内容清洗...")
        filtered_new_emails = AI_Handler.get_conversation_constitutes_for_emails(
            ai_client, emails_to_filter_slow
        )
        print(f"信息：AI 清洗完成，{len(filtered_new_emails)} 封邮件被确认为新对话。")
        emails_to_add_fast.extend(filtered_new_emails)

    # --- 7. (归档) ---
    print(f"信息：开始归档 {len(emails_to_add_fast)} 封最终邮件...")

    new_email_added_count = 0
    addresses_that_were_updated = set()
    processed_email_ids_in_this_run = set(existing_ids)  # (修正：使用 existing_ids 初始化)

    for email in emails_to_add_fast:
        email_id = email.get("id")
        email_type = email.get("type")

        other_party_addresses = []
        if email_type == "received":
            if email.get("sender"):
                other_party_addresses.append(email["sender"])
        elif email_type == "sent":
            receivers = email.get("receiver")
            if isinstance(receivers, list):
                other_party_addresses.extend(receivers)

        if not other_party_addresses:
            print(f"警告：(归档) 邮件 {email_id} 无法确定归档地址，跳过。")
            continue

        for address in set(other_party_addresses):
            if not address: continue

            if address not in all_memory:
                all_memory[address] = {
                    "general_summary": "[新对话：等待AI生成总结]",
                    "style_profile": None,  # (为新对话添加占位符)
                    "emails": []
                }

            all_memory[address]["emails"].append(email)
            addresses_that_were_updated.add(address)

        if email_id and email_id not in processed_email_ids_in_this_run:
            new_email_added_count += 1
            processed_email_ids_in_this_run.add(email_id)
        elif not email_id:
            new_email_added_count += 1

    print(
        f"信息：归档完成。总共添加了 {new_email_added_count} 封新邮件 (分布在 {len(addresses_that_were_updated)} 个对话中)。")

    # --- 8. (排序) ---
    print("信息：正在对所有受影响的对话进行时间排序...")
    # (在AI分析前排序)
    for address in addresses_that_were_updated:
        if address in all_memory and isinstance(all_memory[address], dict):
            email_list = all_memory[address].get("emails", [])
            if email_list:
                try:
                    email_list.sort(key=get_sortable_time)
                except Exception as e:
                    print(f"警告：对话 {address} 排序失败: {e}")

    # --- 9. (总结与口吻分析) (修改点) ---
    if addresses_that_were_updated:
        print(f"信息：检测到 {len(addresses_that_were_updated)} 条对话有更新，准备调用AI分析管道...")

        # (实现点: 只构建需要更新的子集)
        memory_to_update = {}
        for address in addresses_that_were_updated:
            if address in all_memory:  # (安全检查)
                memory_to_update[address] = all_memory[address]

        # --- (AI Pipeline Step 1: 更新总结) ---
        print("  -> (AI Pipeline 1/2) 正在更新对话总结...")
        memory_with_updated_summaries = AI_Handler.get_history_summary_for_conversation(
            ai_client, memory_to_update
        )
        print("  -> AI 总结更新完毕。")

        # --- (AI Pipeline Step 2: 更新口吻) ---
        print("  -> (AI Pipeline 2/2) 正在更新口吻分析...")
        # (将上一步的结果传入，确保口吻分析函数能保留更新后的总结)
        final_updated_conversations = AI_Handler.get_style_profile_for_conversation(
            ai_client, memory_with_updated_summaries
        )
        print("  -> AI 口吻分析更新完毕。")

        # --- (AI Pipeline Step 3: 合并最终结果) ---
        # (final_updated_conversations 包含了 "general_summary", "style_profile", "emails")
        all_memory.update(final_updated_conversations)
        print("信息：AI 分析结果已合并。")

    else:
        print("信息：没有检测到需要更新的对话总结。")
    # --- (修改结束) ---

    # --- 10. (保存) ---
    try:
        with open(memory_file_path, 'w', encoding='utf-8') as f:
            json.dump(all_memory, f, ensure_ascii=False, indent=2, default=datetime_to_json)
        print(f"信息：对话历史已成功保存到 {memory_file_path}")
    except Exception as e:
        print(f"错误：保存对话历史文件失败 ({e})")

    print("//////////////////对话历史维护完成。//////////////////\n")


# --- 周期获取新增邮件并解析处理 ---
def auto_process(mclient, ai_client):
    # 获取邮箱未读邮件
    fetched_in_emails = fetch_unseen_emails(mclient)

    # 获取邮箱发送邮件
    fetched_sent_emails = fetch_sent_emails(mclient)

    # 对邮件分类存储后获取经过总结的有效邮件和发送的邮件列表
    valid_emails, sent_emails = email_classification(ai_client, fetched_in_emails, fetched_sent_emails)

    # 遍历有效邮件查看是否构成对话,若构成则检查对话历史,若存在则完善对话过程,不存在则建立新的对话历史
    maintain_conversation_history(ai_client, valid_emails, sent_emails)

    return valid_emails, sent_emails


# --- 自动循环和停止的包装函数 ---
def start_auto_process_loop(ai_client, stop_event, interval_seconds=600):
    """
    周期性地运行 auto_process 函数，直到 stop_event 被设置。
    """

    while not stop_event.is_set():
        mclient = None
        print(f"\n[{time_module.strftime('%Y-%m-%d %H:%M:%S')}] 开始执行自动流程...")

        try:
            mclient = connect_and_login_email()

            auto_process(mclient, ai_client)

            print(f"[{time_module.strftime('%Y-%m-%d %H:%M:%S')}] 流程执行完毕。")


        except Exception as e:
            print(f"错误：在 auto_process 期间发生意外错误: {e}")

        finally:
            # --- (关键修改 2) ---
            # 3. 无论成功还是失败, 都在循环结束时登出
            if mclient:
                print("  -> 正在从 IMAP 服务器登出...")
                try:
                    mclient.logout()
                except Exception as logout_e:
                    print(f"  -> 警告：登出时发生错误: {logout_e}")
            mclient = None

        print(f"  -> 下次执行将在 {interval_seconds} 秒后...")

        interrupted = stop_event.wait(timeout=interval_seconds)

        if interrupted:
            break

    print("自动处理循环已收到停止信号，即将退出。")


if __name__ == '__main__':
    # 设置间隔时间
    PROCESS_INTERVAL_SECONDS = 60

    # 加载配置
    ai_client = connect_gemini()

    # 在进行对话历史初始化之前，需要进行邮件读取的初始化来构建初始的有效邮件和发送邮件记录
    init_conversation_history(ai_client)

    # 创建 "停止" 信号
    stop_loop_event = threading.Event()

    # 整个自动流程将定时进行，之后需要完善邮件的来源：qq邮箱、Y！メール、Gmail等。

    # 在后台线程中启动循环
    # (我们使用线程，这样主程序就不会被 "while True" 循环卡住)
    process_thread = threading.Thread(
        target=start_auto_process_loop,
        args=(ai_client, stop_loop_event, PROCESS_INTERVAL_SECONDS),
        daemon=True  # 设置为守护线程，这样主程序退出时它也会退出
    )

    # 主程序等待用户输入 "stop"
    print("\n" + "=" * 50)
    print("自动处理程序正在后台运行...")
    print(f"每 {PROCESS_INTERVAL_SECONDS} 秒检查一次邮件。")
    print("在控制台中输入 'stop' (或按 Enter) 来停止程序。")
    print("=" * 50 + "\n")

    process_thread.start()

    try:
        # 阻塞主线程，直到用户输入
        input()
    except EOFError:
        pass  # 在某些非交互式环境中，input()会立即结束

    # 用户输入后，发送 "停止" 信号
    print("正在发送停止信号... (等待当前周期完成)")
    stop_loop_event.set()

    # 等待后台线程完全退出
    # (这会等待 stop_event.wait() 结束, 确保循环优雅地退出)
    process_thread.join(timeout=20)

    print("程序已完全停止。")
