# -*- coding: utf-8 -*-
"""
本地服务器启动脚本
用于运行新股市场情况看板网站
"""
import http.server
import socketserver
import os
import socket

def get_local_ip():
    """获取本机IP地址"""
    try:
        # 创建一个UDP socket来获取本机IP
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "127.0.0.1"

def start_server(port=8080):
    """启动HTTP服务器"""
    try:
        # 确保在正确的目录
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
        
        # 创建服务器
        handler = http.server.SimpleHTTPRequestHandler
        httpd = socketserver.TCPServer(("", port), handler)
        
        # 获取IP地址
        local_ip = get_local_ip()
        
        print("=" * 50)
        print("🚀 新股市场看板服务器")
        print("=" * 50)
        print("服务器启动成功！")
        print(f"本地访问: http://localhost:{port}")
        print(f"局域网访问: http://{local_ip}:{port}")
        print(f"朋友可以用: http://{local_ip}:{port} (需在同一WiFi下)")
        print("按 Ctrl+C 停止服务器")
        print("=" * 50)
        
        # 启动服务器
        httpd.serve_forever()
        
    except KeyboardInterrupt:
        print("\n服务器已停止")
    except OSError as e:
        if "10048" in str(e):
            print(f"❌ 端口 {port} 已被占用！")
            print("解决方法：")
            print("1. 尝试使用其他端口")
            print("2. 或者停止占用该端口的程序")
        else:
            print(f"❌ 服务器启动失败: {e}")

if __name__ == "__main__":
    start_server()