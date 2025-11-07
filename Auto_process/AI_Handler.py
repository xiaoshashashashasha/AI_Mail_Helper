import json
import os
import time

from datetime import datetime
from Auto_process.mail_AutoProcess import VALID_SCORE, CURRENT_DIR, AI_CONFIG
from Auto_process.mail_AutoProcess import TIMEZONE
from Utils.util import datetime_to_json

PROMPT_FILE_PATH = os.path.join(CURRENT_DIR, "../Configs/Prompt_config.json")
JUDGMENT_RECORD_PATH = os.path.join(CURRENT_DIR, "../Info/mail_judgement_record.json")

# ---AI API访问频率限制 ---
SECONDS_BETWEEN_REQUESTS = AI_CONFIG['SECONDS_BETWEEN_REQUESTS']

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
CONVO_PROMPT = prompt_file["CONVERSATION"]
HISTORY_SUMMARY = prompt_file["HISTORY_SUMMARY"]
HISTORY_UPDATE = prompt_file["HISTORY_UPDATE"]
CREATE_STYLE_PROMPT = prompt_file["CREATE_STYLE_PROFILE"]
UPDATE_STYLE_PROMPT = prompt_file["UPDATE_STYLE_PROFILE"]

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


# --- 邮件记录保存 ---
def save_mail_judgment_record(new_records, judgment_type):
    """
    读取现有邮件判断记录，追加新记录，并写入文件，避免覆盖。

    Args:
        new_records (list): 包含AI评分和总结的邮件字典列表。
        judgment_type (string): 该次判断类型
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
    for new_record in new_records:
        new_record["judgment_type"] = judgment_type
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


# --- 邮件分类 ---
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
        judge_list.append(email_data.copy())

        # 若为无效邮件，则在结果中去除总结部分再输出
        if  email_data['score'] < VALID_SCORE:
            email_data.pop('summary',None)

        email_data.pop('judge_time',None)
        result_list.append(email_data)

        time.sleep(SECONDS_BETWEEN_REQUESTS)

    # 将数据结构完备的判断记录存储到../Info/mail_judgement_record.json中
    if judge_list:
        save_mail_judgment_record(judge_list, "classification")

    return result_list


# --- 邮件总结 ---
def get_summary_for_emails(ai_client, emails, model_name="gemini-2.5-flash"):
    """
       对已验证的有效邮件内容进行 AI 总结。

       Args:
           ai_client: 已经初始化的 genai.Client 实例。
           emails: 待处理的邮件字典列表。
           model_name: 使用的模型名称

       Returns:
           list: 包含已更新 summary 字段的邮件字典列表。
       """

    result_list = []
    judge_list = []

    SYSTEM_PROMPT = SUMMARY_PROMPT.get("SYSTEM_PROMPT", "")
    SUMMARY_TASK = SUMMARY_PROMPT.get("SUMMARY_TASK", "")
    RESPONSE_INSTRUCTION = SUMMARY_PROMPT.get("RESPONSE_FORMAT_INSTRUCTION", "")

    print("开始生成有效邮件内容总结")

    for email_data in emails:

        subject = email_data.get('subject', '无主题')
        body = email_data.get('body', '无正文')

        sender_display = (
                email_data.get('sender_name') or
                email_data.get('sender') or
                f"[Root: {email_data.get('sender_root', '未知根域名')}]"
        )

        # --- 1. 构造最终 Prompt ---
        final_prompt = (
                SYSTEM_PROMPT + "\n\n" +
                SUMMARY_TASK + "\n\n" +

                f"邮件主题：{subject}\n" +
                f"邮件正文（仅前1000字）：{body[:1000]}\n\n" +  # 限制长度以节省 token

                RESPONSE_INSTRUCTION
        )


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

            print(f"  AI SUMMARY SUCCESS -> 地址: {sender_display}, 总结: {summary}")

        except Exception as e:
            # 5. 处理 API 失败或 JSON 解析失败
            email_data['summary'] = f"AI处理失败: {e}"
            print(f"  AI SUMMARY FAIL -> 地址: {sender_display}, 错误: {e}")

        # 记录总结处理时间
        email_data['judge_time'] = datetime.now(TIMEZONE).isoformat()
        judge_list.append(email_data.copy())

        # 6. 将处理后的邮件数据添加到结果列表
        email_data.pop('judge_time',None)
        result_list.append(email_data)

        time.sleep(SECONDS_BETWEEN_REQUESTS)

    # 将数据结构完备的判断记录存储到../Info/mail_judgement_record.json中
    if judge_list:
        save_mail_judgment_record(judge_list,"get_summary")

    return result_list


# --- 对话邮件筛选 ---
def get_conversation_constitutes_for_emails(ai_client, emails, model_name="gemini-2.5-flash"):
    """
    [AI-Powered] 使用 AI 进一步筛选邮件，判断其是否构成真实对话（排除系统通知、报告等）。
    此函数 *不* 依赖 classification 的 score，而是进行独立的、更精确的 AI 判断。

    Args:
        ai_client: 已经初始化的 genai.Client 实例。
        emails: 待处理的邮件字典列表。
        model_name: 使用的模型名称

    Returns:
        list: 仅包含被 AI 判断为 "对话型" 的邮件字典列表。
    """

    # result_list: 最终返回的、确认是“对话”的邮件
    result_list = []
    # judge_list: 包含所有处理记录的日志（无论是否被过滤）
    judge_list = []

    # --- 1. 加载 Prompts ---
    try:
        SYSTEM_PROMPT = CONVO_PROMPT.get("SYSTEM_PROMPT", "")
        CONVO_TASK = CONVO_PROMPT.get("CONVO_TASK", "")
        RESPONSE_INSTRUCTION = CONVO_PROMPT.get("RESPONSE_FORMAT_INSTRUCTION", "")
    except KeyError:
        print(f"FATAL: 配置文件 {PROMPT_FILE_PATH} 中缺少 'CONVERSATION' 键。")
        print("警告：由于 Prompt 缺失，将跳过 AI 对话筛选，保留所有邮件。")
        return emails
    except Exception as e:
        print(f"FATAL: 加载 CONVERSATION prompt 失败: {e}")
        return emails

    print("开始进行 AI 对话邮件筛选 (第二阶段)...")

    for email_data in emails:
        # 创建一个副本用于日志记录
        judge_record = email_data.copy()

        subject = email_data.get('subject', '无主题')
        body = email_data.get('body', '无正文')

        sender_display = (
                email_data.get('sender',"未知域名")
        )

        # --- 2. 构造最终 Prompt ---
        final_prompt = (
                SYSTEM_PROMPT + "\n\n" +
                CONVO_TASK + "\n\n" +
                f"邮件主题：{subject}\n" +
                f"邮件正文（仅前1000字）：{body[:1000]}\n\n" +
                RESPONSE_INSTRUCTION
        )

        # 默认为 True
        is_conversation = True
        ai_error_note = None
        judgment_reason = "N/A"

        try:
            # --- 3. 调用 Gemini API ---
            response = retry_gemini_call(
                ai_client.models.generate_content,
                model=model_name,
                contents=final_prompt,
                config={
                    "response_mime_type": "application/json"
                }
            )

            # --- 4. 解析 JSON 结果 ---
            result = json.loads(response.text)

            # (安全地获取布尔值)
            is_conversation_raw = result.get('is_conversation', True)
            is_conversation = str(is_conversation_raw).lower() == 'true'

            judgment_reason = result.get('reason', 'AI未提供理由')

            if is_conversation:
                print(f"  AI CONVO_CHECK -> (保留) 地址: {sender_display}, 理由: {judgment_reason}")
            else:
                print(f"  AI CONVO_CHECK -> (过滤) 地址: {sender_display}, 理由: {judgment_reason}")

        except Exception as e:
            # --- 5. 处理 API 失败 ---
            is_conversation = True  # 安全默认值: 保留
            ai_error_note = f"AI判断对话失败: {e}"
            judgment_reason = "AI判断失败"
            print(f"  AI CONVO_FAIL -> (保留) 地址: {sender_display}, 错误: {e}")

        # --- 6. 记录判断日志 ---
        judge_record['judge_time'] = datetime.now(TIMEZONE).isoformat()
        judge_record['is_conversation_judgment'] = is_conversation  # 记录AI的判断
        judge_record['judgment_reason'] = judgment_reason  # 记录AI的理由
        if ai_error_note:
            judge_record['ai_error'] = ai_error_note  # 记录错误

        # 将日志副本添加到 judge_list
        judge_list.append(judge_record)

        # --- 7. 构建最终返回列表 ---
        if is_conversation:
            # 添加原始邮件数据
            result_list.append(email_data)

        time.sleep(SECONDS_BETWEEN_REQUESTS)

    # --- 8. 保存判断记录 ---
    if judge_list:
        save_mail_judgment_record(judge_list, "conversation_check")

    print(f"AI 对话筛选完成： {len(emails)} 封邮件中，{len(result_list)} 封被保留为对话邮件。")

    return result_list


# --- 对话历史总结 ---
def get_history_summary_for_conversation(ai_client, memory_dict, model_name="gemini-2.5-flash"):
    """
    (最终版) 遍历 *所有* 对话历史，并为 *每一个* 历史生成或更新AI总体总结。
    (修正) 此版本会 *保留* 传入的 "style_profile" 键 (如果存在)。

    Args:
        ai_client: 已经初始化的 genai.Client 实例。
        memory_dict (dict):
            - 旧格式: {"address": [email_list]}
            - 或 新格式: {"address": {"general_summary": "...", "style_profile": ..., "emails": [...]}}
        model_name: 使用的模型名称

    Returns:
        dict: *始终*返回新格式 {"address": {"general_summary": "...", "style_profile": ..., "emails": [...]}}
    """

    print(f"开始为 {len(memory_dict)} 条对话历史生成/更新总体总结...")

    new_memory_structure = {}
    total_conversations = len(memory_dict)
    current_convo_num = 0

    for address, value in memory_dict.items():
        current_convo_num += 1
        print(f"  [总结 {current_convo_num}/{total_conversations}] 正在处理: {address}")

        # --- 1. (格式检测) ---
        email_list = []
        old_summary = None
        old_style_profile = None  # <-- (新增) 默认为空

        if isinstance(value, list):
            # (检测到旧格式: 来自 init)
            email_list = value
            # (old_style_profile 保持 None)
        elif isinstance(value, dict):
            # (检测到新格式: 来自 maintain)
            email_list = value.get("emails", [])
            old_summary = value.get("general_summary", None)
            old_style_profile = value.get("style_profile", None)  # <-- (新增) 获取已有的口吻
        else:
            print(f"    -> 警告: {address} 的数据格式无法识别，跳过。")
            continue

        if not email_list:
            print("    -> 空对话，跳过。")
            new_memory_structure[address] = {
                "general_summary": "空对话历史。",
                "style_profile": old_style_profile, # <-- (新增) 保留 (即使是 None)
                "emails": []
            }
            continue

        # --- 2. 构造 Prompt 输入 (对话摘要) ---
        digest_lines = []
        for email in email_list:  # (假设 email_list 已排序)
            speaker = "[我]" if email.get("type") == "sent" else "[对方]"
            summary = email.get("summary", "无总结")
            subject = email.get("subject", "无主题")
            digest_lines.append(f"{speaker} (主题: {subject}): {summary}")
        conversation_digest = "\n".join(digest_lines)

        # --- 3. (关键: 智能选择 Prompt) ---
        if old_summary and "AI处理失败" not in old_summary:
            # --- (A) 使用 UPDATE 提示词 ---
            print("    -> 检测到旧总结，执行[更新]操作...")
            SYSTEM_PROMPT = HISTORY_UPDATE.get("SYSTEM_PROMPT", "")
            SUMMARY_TASK = HISTORY_UPDATE.get("SUMMARY_TASK", "")
            RESPONSE_INSTRUCTION = HISTORY_UPDATE.get("RESPONSE_FORMAT_INSTRUCTION", "")

            final_prompt = (
                    SYSTEM_PROMPT + "\n\n" +
                    SUMMARY_TASK + "\n\n" +
                    f"【旧的总结】:\n{old_summary}\n\n" +
                    f"【完整的对话摘要 (按时间顺序)】:\n{conversation_digest[:3000]}\n\n" +
                    RESPONSE_INSTRUCTION
            )
        else:
            # --- (B) 使用 CREATE (History) 提示词 ---
            if old_summary:
                print("    -> 旧总结处理失败，执行[重新生成]操作...")
            else:
                print("    -> 未检测到旧总结，执行[创建]操作...")

            SYSTEM_PROMPT = HISTORY_SUMMARY.get("SYSTEM_PROMPT", "")
            SUMMARY_TASK = HISTORY_SUMMARY.get("SUMMARY_TASK", "")
            RESPONSE_INSTRUCTION = HISTORY_SUMMARY.get("RESPONSE_FORMAT_INSTRUCTION", "")

            final_prompt = (
                    SYSTEM_PROMPT + "\n\n" +
                    SUMMARY_TASK + "\n\n" +
                    f"以下是按时间顺序排列的对话摘要:\n{conversation_digest[:3000]}\n\n" +
                    RESPONSE_INSTRUCTION
            )

        # --- 4. 调用 API (try/except 块) ---
        try:
            response = retry_gemini_call(
                ai_client.models.generate_content,
                model=model_name,
                contents=final_prompt,
                config={"response_mime_type": "application/json"}
            )
            result = json.loads(response.text)
            summary = result.get('general_summary', 'AI未提供总体总结')
            print(f"    AI GEN_SUMMARY SUCCESS -> 总结: {summary[:30]}...")

        except Exception as e:
            summary = f"AI处理失败: {e}"
            print(f"    AI GEN_SUMMARY FAIL -> 错误: {e}")

        # --- 5. 构建新结构 ---
        new_memory_structure[address] = {
            "general_summary": summary, # (新生成的总结)
            "style_profile": old_style_profile, # <-- (新增) 保留传入的口吻
            "emails": email_list  # (email_list 是已排序的列表)
        }

        # --- 6. 速率限制 ---
        time.sleep(SECONDS_BETWEEN_REQUESTS)

    # 循环结束
    print(f"信息：新数据结构转换完成 (共 {len(new_memory_structure)} 条对话)。")
    return new_memory_structure


# --- 对话口吻分析 ---
def get_style_profile_for_conversation(ai_client, memory_with_summaries, model_name="gemini-2.5-flash"):
    """
    (修正版) 遍历 *已经包含总结* 的对话历史，并为 *每一个* 历史生成或更新AI口吻分析 (style_profile)。
    (修正) 增强了 JSON 解析的健壮性，以处理扁平(flat)响应。
    (修正) 添加了对 'tone_description' 键的支持。
    """
    print(f"开始为 {len(memory_with_summaries)} 条对话历史生成/更新口吻分析...")

    final_memory_structure = {}
    total_conversations = len(memory_with_summaries)
    current_convo_num = 0

    # --- (修改点 1: 添加 new_key) ---
    default_style_profile = {
        "formality": "未知",
        "tone_description": "未知", # <-- (新增)
        "greeting_template": "无",
        "sign_off_template": "无"
    }
    # --- (修改结束) ---

    for address, value in memory_with_summaries.items():
        current_convo_num += 1
        print(f"  [口吻 {current_convo_num}/{total_conversations}] 正在处理: {address}")

        # --- 1. (格式检测与数据提取) ---
        if not isinstance(value, dict):
            print(f"    -> 警告: {address} 的数据格式不是字典，跳过。")
            continue

        email_list = value.get("emails", [])
        general_summary = value.get("general_summary", "总结丢失")
        old_style_profile = value.get("style_profile", None)

        # --- 2. 构造 Prompt 输入 (口吻分析) ---
        sent_emails = [e for e in email_list if e.get("type") == "sent"]

        if not sent_emails:
            print("    -> 没有 'sent' 邮件，无法分析口吻，跳过。")
            final_memory_structure[address] = {
                "general_summary": general_summary,
                "style_profile": default_style_profile.copy(), # (使用默认值)
                "emails": email_list
            }
            continue

        recent_bodies = [e.get("body", "") for e in sent_emails[-5:]]
        style_digest = "\n\n--- (下一封邮件) ---\n\n".join(recent_bodies)

        # --- 3. (智能选择 Prompt) ---
        if old_style_profile and old_style_profile.get("formality", "未知") != "未知":
            # (A) 使用 UPDATE 提示词
            print("    -> 检测到旧口吻，执行[更新]操作...")
            SYSTEM_PROMPT = UPDATE_STYLE_PROMPT.get("SYSTEM_PROMPT", "")
            SUMMARY_TASK = UPDATE_STYLE_PROMPT.get("SUMMARY_TASK", "")
            RESPONSE_INSTRUCTION = UPDATE_STYLE_PROMPT.get("RESPONSE_FORMAT_INSTRUCTION", "")

            final_prompt = (
                    SYSTEM_PROMPT + "\n\n" +
                    SUMMARY_TASK + "\n\n" +
                    f"【旧的风格分析】:\n{json.dumps(old_style_profile, ensure_ascii=False)}\n\n" +
                    f"【最新的邮件正文 (按时间顺序)】:\n{style_digest[:3000]}\n\n" +
                    RESPONSE_INSTRUCTION
            )
        else:
            # (B) 使用 CREATE 提示词
            if old_style_profile:
                print("    -> 旧口吻无效，执行[重新生成]操作...")
            else:
                print("    -> 未检测到旧口吻，执行[创建]操作...")

            SYSTEM_PROMPT = CREATE_STYLE_PROMPT.get("SYSTEM_PROMPT", "")
            SUMMARY_TASK = CREATE_STYLE_PROMPT.get("SUMMARY_TASK", "")
            RESPONSE_INSTRUCTION = CREATE_STYLE_PROMPT.get("RESPONSE_FORMAT_INSTRUCTION", "")

            final_prompt = (
                    SYSTEM_PROMPT + "\n\n" +
                    SUMMARY_TASK + "\n\n" +
                    f"以下是[我]发送的邮件正文 (按时间顺序):\n{style_digest[:3000]}\n\n" +
                    RESPONSE_INSTRUCTION
            )
        # --- (Prompt 构造结束) ---

        # --- 4. 调用 API (try/except 块) (已修正) ---
        try:
            response = retry_gemini_call(
                ai_client.models.generate_content,
                model=model_name,
                contents=final_prompt,
                config={"response_mime_type": "application/json"}
            )
            result = json.loads(response.text)

            # (修正点: 健壮的解析逻辑)
            style_profile = default_style_profile.copy()  # 先从默认值开始

            if 'style_profile' in result and isinstance(result['style_profile'], dict):
                # (A) 理想情况: AI 遵守了嵌套格式
                print("    -> (解析) AI 遵守了 'style_profile' 嵌套格式。")
                style_profile.update(result['style_profile'])  # .update() 会自动处理新键

            elif 'formality' in result:
                # (B) 备用方案: AI 返回了扁平(flat)格式
                print("    -> (解析) 警告: AI 返回了扁平格式，正在手动构建。")
                style_profile['formality'] = result.get('formality', '未知')
                # --- (修改点 2: 添加 new_key) ---
                style_profile['tone_description'] = result.get('tone_description', '未知') # <-- (新增)
                # --- (修改结束) ---
                style_profile['greeting_template'] = result.get('greeting_template', '无')
                style_profile['sign_off_template'] = result.get('sign_off_template', '无')
            else:
                # (C) 失败情况: AI 返回了无法识别的 JSON
                print("    -> (解析) 警告: AI 未返回 'style_profile' 或 'formality' 键。")
                # (style_profile 保持为 default_style_profile)

            print(f"    AI STYLE SUCCESS -> 格式: {style_profile.get('formality', 'N/A')}")

        except Exception as e:
            style_profile = default_style_profile.copy()
            style_profile["error"] = f"AI处理失败: {e}"
            print(f"    AI STYLE FAIL -> 错误: {e}")
        # --- (修正结束) ---

        # --- 5. (构建 *完整* 结构) ---
        final_memory_structure[address] = {
            "general_summary": general_summary,
            "style_profile": style_profile,
            "emails": email_list
        }

        # --- 6. 速率限制 ---
        time.sleep(SECONDS_BETWEEN_REQUESTS)

    # 循环结束
    print(f"信息：口吻分析转换完成 (共 {len(final_memory_structure)} 条对话)。")
    return final_memory_structure