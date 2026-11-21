import requests
import json

api_key = "sk-54276de9c6ee4af39c5cc94f1748abda"
url = "https://api.deepseek.com/v1/chat/completions"

headers = {
    "Content-Type": "application/json",
    "Authorization": f"Bearer {api_key}"
}

data = {
    "model": "deepseek-chat",
    "messages": [
        {
            "role": "system",
            "content": "你是一名专业的心理治疗师，擅长分析患者的情绪数据并生成详细的心理评估报告。"
        },
        {
            "role": "user",
            "content": "测试消息"
        }
    ],
    "temperature": 0.7,
    "max_tokens": 100
}

try:
    response = requests.post(url, headers=headers, json=data, timeout=10)
    print(f"Response status: {response.status_code}")
    print(f"Response text: {response.text}")
except Exception as e:
    print(f"Error: {str(e)}")
