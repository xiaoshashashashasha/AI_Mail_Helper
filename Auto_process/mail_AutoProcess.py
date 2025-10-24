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
from Utils.util import datetime_to_json, extract_text_from_html, get_address_list_from_header

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
MAIL_CONFIG_FILE = os.path.join(CURRENT_DIR, "../Configs/Setup/mail_config.json")
AI_CONFIG_FILE = os.path.join(CURRENT_DIR,"../Configs/Setup/AI_config.json")
IN_RAWDATA_OUTPUT_PATH = os.path.join(CURRENT_DIR, "../Info/inbox_data.json")
SENT_RAWDATA_OUTPUT_PATH = os.path.join(CURRENT_DIR, "../Info/sentbox_data.json")
SCORE_LIST_PATH = os.path.join(CURRENT_DIR, "../Info/email_sender_score_list.json")
VALID_MAIL_OUTPUT_PATH = os.path.join(CURRENT_DIR, "../Info/valid_emails.json")
INVALID_MAIL_OUTPUT_PATH = os.path.join(CURRENT_DIR, "../Info/invalid_emails.json")
SENT_MAIL_OUTPUT_PATH = os.path.join(CURRENT_DIR, "../Info/sent_emails.json")

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

        to_recipients = []
        for email_addr in to:
            name, root = email_addr.split('@')
            to_recipients.append({'to_name': name, 'to_root': root})

        # 结构化返回信息
        emails.append({
            'type': 'received',
            'id': email_id.decode(),
            'sender_root': sender_root,
            'sender_name': sender_name,
            'to': to_recipients,
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

        to_recipients = []
        for email_addr in to:
            name, root = email_addr.split('@')
            to_recipients.append({'to_name': name, 'to_root': root})

        # 结构化返回信息
        emails.append({
            'type': 'sent',
            'id': email_id.decode(),
            'sender': email_addr,
            'to': to_recipients,
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



    if not (in_emails or sent_emails):
        print("无邮件需要分类")
        return []

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

# --- 根据获取的有效邮件维护对话历史 ---
def maintain_conversation_history(valid_emails):
    print("对话历史维护")
    # 具体实现：
    # 1.根据邮件内容判断是否可能为对话邮件
    # 2.根据往来域名确认对话双方
    # 3.逐一检查邮件内容是否与该交流组的对话内的各对话相关
    # 4.若相关，则归入该对话中；不相关，则在该交流组内创建新的对话


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
    valid_emails = email_classification(ai_client, fetched_in_emails, fetched_sent_emails)

    # 遍历有效邮件查看是否构成对话,若构成则检查对话历史,若存在则完善对话过程,不存在则建立新的对话历史
    maintain_conversation_history(valid_emails)


    return valid_emails

if __name__ == '__main__':

    # 加载配置
    mclient = connect_and_login_email()
    ai_client = connect_gemini()

    auto_process(mclient, ai_client)


    # 整个自动流程将定时进行，之后需要完善邮件的来源：qq邮箱、Y！メール、Gmail等。


    mclient.logout()
