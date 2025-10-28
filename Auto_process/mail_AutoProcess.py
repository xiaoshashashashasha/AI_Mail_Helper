import imaplib
import email
import json
import os
import pytz
import AI_Handler

from datetime import datetime
from email.utils import parseaddr
from email.header import decode_header
from email.utils import parsedate_to_datetime
from google import genai
from Utils.util import datetime_to_json, extract_text_from_html, get_address_list_from_header, archive_email_to_memory

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
    print("对话历史初始化...")

    # --- 1. (约束检查) 仅能在对话历史为空时执行 ---
    if os.path.exists(memory_file_path):
        try:
            # 检查文件大小 > 10 字节 (防止只是一个 "{}")
            if os.path.getsize(memory_file_path) > 10:
                with open(memory_file_path, 'r', encoding='utf-8') as f:
                    content = json.load(f)
                # 检查内容是否为空
                if content:
                    print(f"错误：对话历史文件 {memory_file_path} 已存在且不为空。初始化被终止。")
                    print("如需强制重新初始化，请手动删除该文件。")
                    return
        except Exception as e:
            print(f"警告：对话历史文件 {memory_file_path} 存在但无法解析({e})。为安全起见，初始化被终止。")
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
        return
    except Exception as e:
        print(f"FATAL: 无法读取历史邮件文件: {e}")
        return

    print(f"信息：已加载 {len(all_valid_emails)} 封收信和 {len(all_sent_emails)} 封发信。")

    # --- 3. (统一格式化) 处理 valid_emails (收信) ---
    formatted_valid_emails = []
    if len(all_valid_emails) > 0:
        print(f"开始对 {len(all_valid_emails)} 封历史收信进行格式化...")
        for email in all_valid_emails:
            # (地址清洗) 检查 NO_REPLY
            if any(pattern in email.get("sender_name", "").lower() for pattern in NO_REPLY_PATTERN):
                continue

            # (格式化)
            sender_addr = email.get("sender_name", "unknown") + "@" + email.get("sender_root", "unknown.com")

            if not sender_addr or sender_addr == "unknown@unknown.com":
                continue  # 跳过无法识别的邮件

            email["sender"] = sender_addr
            email.pop("sender_name", None)
            email.pop("sender_root", None)
            email.pop("score", None)

            formatted_valid_emails.append(email)
        print(f"信息：历史收信格式化完成，{len(formatted_valid_emails)} 封邮件有效。")

    # --- 4. (统一格式化) 处理 sent_emails (发信) ---
    formatted_sent_emails = []
    if len(all_sent_emails) > 0:
        for email in all_sent_emails:
            receivers = email.get("receiver")
            if not isinstance(receivers, list) or not receivers:
                continue
            formatted_sent_emails.append(email)

    # --- 5. (慢速通道) 运行 AI 清洗 (针对所有收信) ---
    # 在 init 模式下, *所有* 收信都必须通过 AI 验证

    if formatted_valid_emails:
        print(f"信息：正在提交 {len(formatted_valid_emails)} 封历史收信到 AI 进行内容清洗...")
        print("(这可能需要很长时间，取决于邮件数量...)")

        # (a. 提交给 AI)
        final_valid_emails = AI_Handler.get_conversation_constitutes_for_emails(
            ai_client, formatted_valid_emails
        )
        print(f"信息：AI 清洗完成，{len(final_valid_emails)} 封邮件被确认为对话。")
    else:
        final_valid_emails = []

    # --- 6. (归档) ---
    # (发信总是被视为对话，无需 AI 过滤)
    all_emails_to_archive = formatted_sent_emails + final_valid_emails

    if not all_emails_to_archive:
        print("警告：没有可用于归档的邮件。初始化终止。")
        return

    print(f"信息：开始归档 {len(all_emails_to_archive)} 封最终邮件...")

    all_memory = {}  # (初始化：从空字典开始)
    new_email_added_count = 0
    processed_ids_this_run = set()

    for email in all_emails_to_archive:
        new_email_added_count += archive_email_to_memory(
            email, all_memory, processed_ids_this_run
        )

    print(f"信息：归档完成。总共添加了 {new_email_added_count} 封邮件到 {len(all_memory)} 条对话中。")

    # --- 7. (排序与保存) ---
    print("信息：正在对所有对话进行时间排序...")
    for email_list in all_memory.values():
        if not isinstance(email_list, list): continue
        try:
            email_list.sort(key=lambda x: x.get("sent_time", "1970-01-01T00:00:00Z"))
        except Exception as e:
            print(f"警告：对话 {e} 排序失败。")

    try:
        with open(memory_file_path, 'w', encoding='utf-8') as f:
            json.dump(all_memory, f, ensure_ascii=False, indent=2)
        print(f"信息：对话历史已成功初始化并保存到 {memory_file_path}")
    except Exception as e:
        print(f"错误：保存对话历史文件失败 ({e})")

    print("对话历史初始化完成。")


# --- 根据获取的有效邮件维护对话历史 ---
def maintain_conversation_history(ai_client, valid_emails, sent_emails, memory_file_path=CONVERSATION_MEMORY_PATH):
    # 步骤：
    # 1. 读取对话历史 (all_memory, existing_ids, existing_addresses)
    # 2. (统一格式化) 对 valid_emails 进行地址清洗和数据格式化
    # 3. (统一格式化) 筛选 sent_emails
    # 4. 将格式化后的邮件分流 (快/慢通道)
    # 5. (慢速通道) 运行 AI 清洗
    # 6. (归档) 将 快通道 + AI清洗后的慢通道 邮件统一归档
    # 7. 排序并保存

    if not (valid_emails or sent_emails):
        print("无可用于对话历史维护的邮件")
        return

    # --- 1. 读取对话历史记录 ---
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
    for email_list in all_memory.values():
        if not isinstance(email_list, list): continue
        for email in email_list:
            if email and email.get("id"):
                existing_ids.add(email.get("id"))

    existing_addresses = set(all_memory.keys())
    print(f"信息：已加载 {len(existing_ids)} 个邮件ID 和 {len(existing_addresses)} 条已有对话。")

    # --- 2. (统一格式化) 处理 valid_emails (收信) ---
    formatted_valid_emails = []
    if len(valid_emails) > 0:
        print(f"开始对 {len(valid_emails)} 封收信进行格式化...")
        for email in valid_emails:
            # (去重) 检查是否已在 memory 中
            if email.get("id") and email.get("id") in existing_ids:
                continue

            # (地址清洗) 检查 NO_REPLY
            if any(pattern in email.get("sender_name", "").lower() for pattern in NO_REPLY_PATTERN):
                continue

            # (格式化)
            sender_addr = email.get("sender_name", "unknown") + "@" + email.get("sender_root", "unknown.com")

            if not sender_addr or sender_addr == "unknown@unknown.com":
                print(f"警告：(收信) 邮件 {email.get('id')} 无法确定发件人地址，跳过。")
                continue

            email["sender"] = sender_addr
            email.pop("sender_name", None)
            email.pop("sender_root", None)
            email.pop("score", None)

            formatted_valid_emails.append(email)
        print(f"信息：收信格式化完成，{len(formatted_valid_emails)} 封邮件有效。")

    # --- 3. (统一格式化) 处理 sent_emails (发信) ---
    formatted_sent_emails = []
    if len(sent_emails) > 0:
        for email in sent_emails:
            # (去重) 检查是否已在 memory 中
            if email.get("id") and email.get("id") in existing_ids:
                continue

            receivers = email.get("receiver")
            if not isinstance(receivers, list) or not receivers:
                print(f"警告：(发信) 邮件 {email.get('id')} 缺少收件人，跳过。")
                continue

            formatted_sent_emails.append(email)

    # --- 4. 将格式化后的邮件分流 ---
    emails_to_add_fast = []  # 快速通道
    emails_to_filter_slow = []  # 慢速通道

    # (分流收信)
    for email in formatted_valid_emails:
        if email["sender"] in existing_addresses:
            emails_to_add_fast.append(email)
        else:
            emails_to_filter_slow.append(email)

    # (分流发信 - 总是快速)
    emails_to_add_fast.extend(formatted_sent_emails)

    print(f"信息：邮件分流完成。快速通道: {len(emails_to_add_fast)} 封，慢速(AI)通道: {len(emails_to_filter_slow)} 封。")

    # --- 5. (慢速通道) 运行 AI 清洗 ---
    if emails_to_filter_slow:
        print(f"信息：正在提交 {len(emails_to_filter_slow)} 封新邮件到 AI 进行内容清洗...")

        # (a. 提交给 AI)
        filtered_new_emails = AI_Handler.get_conversation_constitutes_for_emails(
            ai_client, emails_to_filter_slow
        )

        print(f"信息：AI 清洗完成，{len(filtered_new_emails)} 封邮件被确认为新对话。")

        # (b. 将 AI 过滤后的结果添加到 "快速通道" 准备归档)
        emails_to_add_fast.extend(filtered_new_emails)

    # --- 6. (归档) ---
    print(f"信息：开始归档 {len(emails_to_add_fast)} 封最终邮件...")
    new_email_added_count = 0
    # (用于处理 "sent" 邮件只被计数一次)
    processed_ids_this_run = set()

    for email in emails_to_add_fast:
        new_email_added_count += archive_email_to_memory(
            email, all_memory, processed_ids_this_run
        )

    print(f"信息：归档完成。总共添加了 {new_email_added_count} 封新邮件。")

    # --- 7. (排序与保存) ---
    print("信息：正在对所有对话进行时间排序...")
    for email_list in all_memory.values():
        if not isinstance(email_list, list): continue
        try:
            email_list.sort(key=lambda x: x.get("sent_time", "1970-01-01T00:00:00Z"))
        except Exception as e:
            print(f"警告：对话 {e} 排序失败。")

    try:
        with open(memory_file_path, 'w', encoding='utf-8') as f:
            json.dump(all_memory, f, ensure_ascii=False, indent=2)
        print(f"信息：对话历史已成功保存到 {memory_file_path}")
    except Exception as e:
        print(f"错误：保存对话历史文件失败 ({e})")

    print("对话历史维护完成。")


# --- 周期获取新增邮件并解析处理 ---
def auto_process(mclient, ai_client):
    # 获取邮箱未读邮件
    fetched_in_emails = fetch_unseen_emails(mclient)
    # 获取邮箱发送邮件
    fetched_sent_emails = fetch_sent_emails(mclient)

    # 后续自动步骤：
    # 识别、总结与对话的记忆
    #   对返回的邮件信息根据域名黑名单进行初筛，保留的邮件交由AI进行总结并甄别，
    #   识别分类-将无效邮件数据存储到json中，无法识别的列为无法识别，交由AI进行判断
    #   根据AI的评分维护邮件地址评价文件,
    #   对有效邮件的type类型进行映射获取邮件属性，完善其数据结构并存储到json中，
    #   总结后的完整邮件信息将用于维护对话记忆表。
    valid_emails, sent_emails = email_classification(ai_client, fetched_in_emails, fetched_sent_emails)

    # 遍历有效邮件查看是否构成对话,若构成则检查对话历史,若存在则完善对话过程,不存在则建立新的对话历史
    maintain_conversation_history(ai_client, valid_emails, sent_emails)


    return valid_emails, sent_emails

if __name__ == '__main__':

    # 加载配置
    mclient = connect_and_login_email()
    ai_client = connect_gemini()

    init_conversation_history(ai_client)
    auto_process(mclient, ai_client)


    # 整个自动流程将定时进行，之后需要完善邮件的来源：qq邮箱、Y！メール、Gmail等。


    mclient.logout()
