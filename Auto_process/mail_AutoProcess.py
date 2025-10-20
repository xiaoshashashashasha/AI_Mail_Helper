import imaplib
import email
import json
import os
import pytz
import AI_Handler
from email.utils import parseaddr
from email.header import decode_header
from email.utils import parsedate_to_datetime
from google import genai
from Utils.util import datetime_to_json



CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
MAIL_CONFIG_FILE = os.path.join(CURRENT_DIR, "../Setup/mail_config.json")
AI_CONFIG_FILE = os.path.join(CURRENT_DIR,"../Setup/AI_config.json")
RAW_OUTPUT_PATH = os.path.join(CURRENT_DIR, "../Info/inbox_data.json")
SCORE_LIST_PATH = os.path.join(CURRENT_DIR, "../Info/email_sender_score_list.json")

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


# --- 读取邮件,结构化并保存为原始数据 ---
def fetch_unseen_emails(mclient, json_file_path=RAW_OUTPUT_PATH):
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


        sender = msg['From']


        body = ""

        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                cdispo = str(part.get('Content-Disposition'))

                if ctype == 'text/plain' and 'attachment' not in cdispo:
                    body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    break
        else:
            if msg.get_content_type() == 'text/plain':
                body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
        # ----------------------------------------------------
        # 转换为指定时区时间
        if sent_time:
            sent_time_jst = sent_time.astimezone(TIMEZONE)
        else:
            sent_time_jst = None

        # 解析sender
        sender_root = ''
        sender_name = ''

        if sender:
            display_name, email_addr = parseaddr(sender)
            sender_name, sender_root = email_addr.split('@')


        # 结构化返回信息
        emails.append({
            'id': email_id.decode(),
            'sender_root': sender_root,
            'sender_name': sender_name,
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
        all_emails.sort(key=lambda email: email['sent_time'])

        # 3. 写入完整合并后的数据 (覆盖旧文件)
        with open(json_file_path, 'w', encoding='utf-8') as f:
            json.dump(all_emails, f, indent=4, ensure_ascii=False, default=datetime_to_json)
        print(f"成功提取 {len(emails)} 封邮件，并写入到 {json_file_path}")
    else:
        print("没有发现新的未读邮件。")

    return emails

# --- 对邮件分类并存储 ---
"""获取评分表，根据匹配到的分数进行第一次分类，无匹配项的则将其交由AI进行评分（摘要和正文内容），随后根据该发件地址对分数列表进行维护"""
def email_classification(ai_client, emails, invalid_output_path="../Info/invalid_emails.json", valid_output_path="../Info/valid_emails.json"):
    valid_emails = []
    invalid_emails = []
    uncertain_emails = []

    if not emails:
        print("无邮件需要分类")
        return []

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
    for email in emails:
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
            email["ai_score"] = score
            valid_emails.append(email)

        else :
            email["ai_score"] = score
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
        all_invalid_emails.sort(key=lambda email: email['sent_time'])

        # 3. 覆盖写入完整列表
        with open(invalid_output_path, 'w', encoding='utf-8') as f:
            json.dump(all_invalid_emails, f, indent=4, ensure_ascii=False, default=datetime_to_json)
            print(f"SUCCESS: {len(invalid_emails)} 封邮件被标记为无效邮件，写入 {invalid_output_path}")


    # 有效邮件完善其数据结构,后续用于对话记忆等功能
    if len(valid_emails) > 0:
        # 对于没有总结的，交由AI进行总结
        need_summarize_list = []
        for email in valid_emails:
            if not email["summary"]:
                need_summarize_list.append(email)

        if len(need_summarize_list) > 0:
            result_list = AI_Handler.get_summary_for_valid_emails(ai_client, need_summarize_list, MODEL_NAME)
            valid_emails.extend(result_list)
            valid_emails.sort(key=lambda email: email['sent_time'])

        # 存储为有效邮件
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
        all_valid_emails.sort(key=lambda email: email['sent_time'])

        # 3. 覆盖写入完整列表
        with open(valid_output_path, 'w', encoding='utf-8') as f:
            json.dump(all_valid_emails, f, indent=4, ensure_ascii=False, default=datetime_to_json)
            print(f"SUCCESS: {len(valid_emails)} 封邮件被标记为有效邮件，将在处理后写入 {valid_output_path},并用于记忆构成")


    return valid_emails


if __name__ == '__main__':

    # 加载配置
    mclient = connect_and_login_email()
    ai_client = connect_gemini()

    # 获取邮箱未读邮件
    fetched_emails = fetch_unseen_emails(mclient, json_file_path='../Info/inbox_data.json')

    # 后续自动步骤：
    # 识别、总结与对话的记忆
    #   对返回的邮件信息根据域名黑名单进行初筛，保留的邮件交由AI进行总结并甄别，
    #   识别分类-将无效邮件数据存储到json中，无法识别的列为无法识别，交由AI进行判断
    #   根据AI的评分维护邮件地址评价文件,
    #   对有效邮件的type类型进行映射获取邮件属性，完善其数据结构并存储到json中，
    #   总结后的完整邮件信息将用于维护对话记忆表。
    valid_emails = email_classification(ai_client, fetched_emails)

    # 遍历有效邮件查看是否构成对话,若构成则检查对话历史,若存在则完善对话过程,不存在则建立新的对话历史


    # 整个自动流程将定时进行，之后需要完善邮件的来源：qq邮箱、Y！メール、Gmail等。


    mclient.logout()
