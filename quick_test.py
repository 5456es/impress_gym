import requests
import time
import os
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
time.sleep(2)

print("测试 API 端点...")

# 1. 健康检查
print("\n1. 健康检查:")
resp = requests.get("http://127.0.0.1:5011/api/health")
print(f"状态码: {resp.status_code}")
print(f"响应: {resp.text}")

# 2. 连接测试
print("\n2. 连接测试:")
resp = requests.post("http://127.0.0.1:5011/api/connect")
print(f"状态码: {resp.status_code}")
print(f"响应: {resp.text}")

# 3. 获取演示文稿信息
print("\n3. 演示文稿信息:")
resp = requests.get("http://127.0.0.1:5011/api/presentation/info")
print(f"状态码: {resp.status_code}")
print(f"响应: {resp.text}")