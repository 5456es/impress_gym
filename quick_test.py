import requests
import time
import os
import random
# 清除代理环境变量

proxy_vars = ['http_proxy', 'HTTP_PROXY', 'https_proxy', 'HTTPS_PROXY', 
              'ftp_proxy', 'FTP_PROXY', 'all_proxy', 'ALL_PROXY']

for var in proxy_vars:
    if var in os.environ:
        del os.environ[var]

# 设置 no_proxy 以确保本地连接不使用代理
os.environ['no_proxy'] = 'localhost,127.0.0.1,::1'
os.environ['NO_PROXY'] = 'localhost,127.0.0.1,::1'

# 等待一下确保服务完全启动

# print("测试 API 端点...")

# params = {
#     "include_formatting": "true"  # 或 "false"
# }
# # /api/slide/current	
# response = requests.post("http://localhost:5011/api/slide/new", json={})
# print("Response from /api/slide/add-text:", response.status_code, response.text)



# response = requests.get("http://localhost:5011/api/slide/current", json={})
# response=requests.post("http://localhost:5011/api/slide/new", json={})



# print("Response from /api/slide/current:", response.status_code, response.text)
# select box
### set up boxes

textboxes=["123", "456", "789"]


for text in textboxes:
    response = requests.post("http://localhost:5011/api/slide/add-text", json={
        "text": text,
        "slide_index":2,
        "x": random.randint(1000, 2000),
        "y": random.randint(1000, 2000),
    })
    print(f"Response from /api/slide/add-text for {text}:", response.status_code, response.text)