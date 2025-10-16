from datetime import datetime



# --- 辅助函数：将 datetime 对象转换为可序列化的字符串 ---
# json 模块不能直接序列化 datetime 对象，需要转换
def datetime_to_json(obj):
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError(f"Object of type {type(obj).__name__} is not JSON serializable")



