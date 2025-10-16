import json
import time

PROMPT_FILE_PATH = "../Setup/Prompt_config.json"

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



# --- 辅助函数：重试机制 (用于处理 API 错误) ---
def retry_gemini_call(func, *args, max_retries=3, delay=5, **kwargs):
    """为 Gemini API 调用添加指数退避重试机制"""
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            if attempt < max_retries - 1:
                # 假设 e 包含 API 错误信息
                print(f"警告：Gemini API 调用失败 ({e})，将在 {delay} 秒后重试... (尝试 {attempt + 1}/{max_retries})")
                time.sleep(delay)
                delay *= 2
            else:
                print(f"FATAL: Gemini API 多次重试失败，跳过此邮件。错误: {e}")
                # 必须重新抛出异常，让调用者知道失败
                raise

def get_score_for_uncertain_emails(ai_client, uncertain_emails, model_name="gemini-2.5-flash"):
    """
    对未分类的邮件进行 AI 评分和总结。

    Args:
        ai_client: 已经初始化的 genai.Client 实例。
        uncertain_emails: 待处理的邮件字典列表。
        model_name: 使用的模型名称

    Returns:
        list: 包含 (email_data_dict, score) 元组的列表。
    """

    result_list = []

    SYSTEM_PROMPT = CLASSIFICATION_PROMPT["SYSTEM_PROMPT"]
    CLASSIFY_TASK = CLASSIFICATION_PROMPT["CLASSIFY_TASK"]
    SCORES_MAPPING = CLASSIFICATION_PROMPT["SCORES"]
    RESPONSE_INSTRUCTION = CLASSIFICATION_PROMPT["RESPONSE_FORMAT_INSTRUCTION"]

    # 将评分映射转换为 AI 可读的字符串格式
    SCORES_STR = "\n".join([f"- {k}: {v}分" for k, v in SCORES_MAPPING.items()])

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

        score = 0  # 默认评分为 0
        summary = "AI处理失败或未总结"

        try:
            # 2. 调用 Gemini API (假设 MODEL_NAME 是从外部传入的或已定义)
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

            # 提取评分 (确保它是整数)
            score = int(result.get('score', 0))
            summary = result.get('summary', '未总结')

            # 4. 更新邮件数据字典 (用于返回和后续处理)
            email_data['ai_score'] = score
            email_data['summary'] = summary

            print(f"  AI SUCCESS -> 地址: {email_data['sender_name']}, 分数: {score}, 总结: {summary}")

        except Exception as e:
            # 5. 处理 API 失败或 JSON 解析失败
            email_data['ai_score'] = 0  # 评分失败，给予最低分（或 0 分）
            email_data['summary'] = f"AI处理失败: {e}"
            print(f"  AI FAIL -> 地址: {email_data['sender_name']}, 错误: {e}")

        # 6. 将邮件数据和评分添加到结果列表 (无论成功失败，都需要返回进行后续分类)
        result_list.append((email_data, email_data['ai_score']))

    return result_list