import json
import time

from datetime import datetime
from Auto_process.mail_AutoProcess import VALID_SCORE
from Auto_process.mail_AutoProcess import TIMEZONE
from Utils.util import datetime_to_json

PROMPT_FILE_PATH = "../Setup/Prompt_config.json"
JUDGMENT_RECORD_PATH = "../Info/mail_judgement_record.json"

try:
    with open(PROMPT_FILE_PATH, 'r', encoding='utf-8') as f:
        prompt_file = json.load(f)
except FileNotFoundError:
    print(f"错误：找不到配置文件 {PROMPT_FILE_PATH}，请检查路径。")
    score_list = {}
except json.JSONDecodeError:
    print(f"错误：配置文件 {PROMPT_FILE_PATH} 格式不正确。")
    score_list = {}

CLASSIFICATION_PROMPT = prompt_file["CLASSIFICATION"]
SUMMARY_PROMPT = prompt_file["SUMMARY"]


# --- 辅助函数：重试机制 (用于处理 API 错误) ---
def retry_gemini_call(func, *args, max_retries=3, delay=5, **kwargs):
    """为 Gemini API 调用添加指数退避重试机制"""
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            if attempt < max_retries - 1:
                print(f"警告：Gemini API 调用失败 ({e})，将在 {delay} 秒后重试... (尝试 {attempt + 1}/{max_retries})")
                time.sleep(delay)
                delay *= 2
            else:
                print(f"FATAL: Gemini API 多次重试失败，跳过此邮件。错误: {e}")
                raise


# --- 邮件记录保存函数 ---
def save_mail_judgment_record(new_records):
    """
    读取现有邮件判断记录，追加新记录，并写入文件，避免覆盖。

    Args:
        new_records (list): 包含AI评分和总结的邮件字典列表。
    """

    # 1. 读取现有记录
    try:
        with open(JUDGMENT_RECORD_PATH, 'r', encoding='utf-8') as f:
            existing_records = json.load(f)
        if not isinstance(existing_records, list):
            # 如果文件内容不是列表，则视为无效，创建空列表
            print(f"警告：历史记录文件 {JUDGMENT_RECORD_PATH} 内容格式不是列表，将清除旧记录并使用新记录。")
            existing_records = []
    except FileNotFoundError:
        print(f"提示：找不到历史记录文件 {JUDGMENT_RECORD_PATH}，将创建新文件。")
        existing_records = []
    except json.JSONDecodeError:
        print(f"警告：历史记录文件 {JUDGMENT_RECORD_PATH} 格式不正确（非有效 JSON），将清除旧记录并使用新记录。")
        existing_records = []
    except Exception as e:
        print(f"警告：读取历史记录文件时发生意外错误 ({e})，将清除旧记录并使用新记录。")
        existing_records = []

    # 2. 合并记录
    combined_records = existing_records + new_records

    # 3. 写入文件
    try:
        with open(JUDGMENT_RECORD_PATH, 'w', encoding='utf-8') as f:
            # ensure_ascii=False 确保中文能正确写入 JSON 文件
            # indent=4 使文件格式更易读
            json.dump(combined_records, f, ensure_ascii=False, indent=4,default=datetime_to_json)
        print(f"信息：成功将 {len(new_records)} 条 AI 判断记录追加并写入文件 {JUDGMENT_RECORD_PATH}。")
    except IOError as e:
        print(f"错误：写入文件 {JUDGMENT_RECORD_PATH} 失败: {e}")


# --- 邮件分类函数 ---
def get_score_for_uncertain_emails(ai_client, uncertain_emails, model_name="gemini-2.5-flash"):
    """
    对未分类的邮件进行 AI 评分和总结，分类记录会保存到JSON中。

    Args:
        ai_client: 已经初始化的 genai.Client 实例。
        uncertain_emails: 待处理的邮件字典列表。
        model_name: 使用的模型名称

    Returns:
        list: 包含 email_data_dict邮件字典的列表。
    """

    result_list = []
    judge_list = []

    SYSTEM_PROMPT = CLASSIFICATION_PROMPT["SYSTEM_PROMPT"]
    CLASSIFY_TASK = CLASSIFICATION_PROMPT["CLASSIFY_TASK"]
    SCORES_MAPPING = CLASSIFICATION_PROMPT["SCORES"]
    RESPONSE_INSTRUCTION = CLASSIFICATION_PROMPT["RESPONSE_FORMAT_INSTRUCTION"]

    # 将评分映射转换为 AI 可读的字符串格式
    SCORES_STR = "\n".join([f"- {k}: {v}分" for k, v in SCORES_MAPPING.items()])

    print("开始对未分类邮件进行分类")

    for email_data in uncertain_emails:
        subject = email_data['subject']
        body = email_data['body']

        # --- 1. 构造最终 Prompt ---
        final_prompt = (
                SYSTEM_PROMPT + "\n\n" +
                CLASSIFY_TASK.format(scores=SCORES_STR) + "\n\n" +

                f"邮件主题：{subject}\n" +
                f"邮件正文（仅前1000字）：{body[:1000]}\n\n" +  # 限制长度以节省 token

                RESPONSE_INSTRUCTION
        )


        try:
            # 2. 调用 Gemini API
            response = retry_gemini_call(
                ai_client.models.generate_content,
                model=model_name,  # 确保 MODEL_NAME 变量可用
                contents=final_prompt,
                config={
                    "response_mime_type": "application/json"
                    # 这里省略 response_schema，因为 prompt 已经严格要求了 JSON 格式
                }
            )

            # 3. 解析 JSON 结果
            result = json.loads(response.text)

            # 提取评分
            score = int(result.get('score', 5))
            summary = result.get('summary', '未总结')

            # 4. 更新邮件数据字典 (用于返回和后续处理)
            email_data['score'] = score
            email_data['summary'] = summary

            print(f"  AI SUCCESS -> 地址: {email_data['sender_name']}, 分数: {score}, 总结: {summary}")

        except Exception as e:
            # 5. 处理 API 失败或 JSON 解析失败
            email_data['score'] = 5  # 评分失败，给予最高分
            email_data['summary'] = f"AI处理失败: {e}"
            print(f"  AI FAIL -> 地址: {email_data['sender_name']}, 错误: {e}")

        # 6. 将邮件数据和评分添加到结果列表和判断列表中
        email_data['judge_time'] = datetime.now(TIMEZONE).isoformat()
        judge_list.append(email_data)

        # 若为无效邮件，则在结果中去除总结部分再输出
        if  email_data['score'] < VALID_SCORE:
            email_data.pop('summary',None)

        email_data.pop('judge_time',None)
        result_list.append(email_data)

    # 将数据结构完备的判断记录存储到../Info/mail_judgement_record.json中
    if judge_list:
        save_mail_judgment_record(judge_list)

    return result_list


# --- 有效邮件总结 ---
def get_summary_for_valid_emails(ai_client, valid_emails, model_name="gemini-2.5-flash"):
    """
       对已验证的有效邮件内容进行 AI 总结。

       Args:
           ai_client: 已经初始化的 genai.Client 实例。
           valid_emails: 待处理的邮件字典列表。
           model_name: 使用的模型名称

       Returns:
           list: 包含已更新 summary 和 summary_time 字段的邮件字典列表。
       """

    result_list = []
    judge_list = []

    SYSTEM_PROMPT = SUMMARY_PROMPT.get("SYSTEM_PROMPT", "")
    SUMMARY_TASK = SUMMARY_PROMPT.get("SUMMARY_TASK", "")
    RESPONSE_INSTRUCTION = SUMMARY_PROMPT.get("RESPONSE_FORMAT_INSTRUCTION", "")

    print("开始生成有效邮件内容总结")

    for email_data in valid_emails:

        subject = email_data.get('subject', '无主题')
        body = email_data.get('body', '无正文')

        # --- 1. 构造最终 Prompt ---
        final_prompt = (
                SYSTEM_PROMPT + "\n\n" +
                SUMMARY_TASK + "\n\n" +

                f"邮件主题：{subject}\n" +
                f"邮件正文（仅前1000字）：{body[:1000]}\n\n" +  # 限制长度以节省 token

                RESPONSE_INSTRUCTION
        )

        summary = "AI总结失败或未总结"

        try:
            # 2. 调用 Gemini API
            response = retry_gemini_call(
                ai_client.models.generate_content,
                model=model_name,
                contents=final_prompt,
                config={
                    "response_mime_type": "application/json"
                }
            )

            # 3. 解析 JSON 结果
            result = json.loads(response.text)

            # 提取总结 (根据 prompt 结构，这里直接提取 'summary' 字段)
            summary = result.get('summary', 'AI未提供总结')

            # 4. 更新邮件数据字典
            email_data['summary'] = summary

            print(f"  AI SUMMARY SUCCESS -> 地址: {email_data.get('sender_name', '未知发送者')}, 总结: {summary}")

        except Exception as e:
            # 5. 处理 API 失败或 JSON 解析失败
            email_data['summary'] = f"AI处理失败: {e}"
            print(f"  AI SUMMARY FAIL -> 地址: {email_data.get('sender_name', '未知发送者')}, 错误: {e}")

        # 记录总结处理时间
        email_data['judge_time'] = datetime.now(TIMEZONE).isoformat()
        judge_list.append(email_data)

        # 6. 将处理后的邮件数据添加到结果列表
        email_data.pop('judge_time',None)
        result_list.append(email_data)

    # 将数据结构完备的判断记录存储到../Info/mail_judgement_record.json中
    if judge_list:
        save_mail_judgment_record(judge_list)

    return result_list