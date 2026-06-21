#!/usr/bin/env python3
"""
销售分析网站启动脚本
"""

import http.server
import socketserver
import json
import os
import sys

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

PORT = 8002

class AnalysisHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory='static', **kwargs)

    def do_POST(self):
        """处理上传请求"""
        if self.path == '/api/upload':
            try:
                # 使用multipart解析
                from multipart.multipart import parse_options
                import io

                content_length = int(self.headers['Content-Length'])
                body = self.rfile.read(content_length)

                # 手动解析multipart数据（简化版）
                boundary = self.headers['Content-Type'].split('boundary=')[1]
                parts = body.split(f'--{boundary}'.encode())

                file_data = None
                file_name = 'unknown.xls'

                for part in parts:
                    if b'Content-Disposition: form-data' in part and b'filename=' in part:
                        # 提取文件名
                        for line in part.split(b'\r\n'):
                            if b'filename=' in line:
                                file_name = line.split(b'"')[1].decode('utf-8')
                        # 提取文件数据
                        data_start = part.find(b'\r\n\r\n') + 4
                        file_data = part[data_start:].rstrip(b'\r\n')
                        break

                if not file_data:
                    raise ValueError("No file data found in request")

                # 保存到临时文件
                os.makedirs('uploads', exist_ok=True)
                temp_path = f'uploads/upload_{file_name}'
                with open(temp_path, 'wb') as f:
                    f.write(file_data)

                print(f"[UPLOAD] 文件已保存: {temp_path} ({len(file_data)} bytes)", flush=True)

                # 调用分析函数
                from analyzer import analyze_sales_data
                result = analyze_sales_data(temp_path)

                # 返回结果
                response = {
                    'id': 'test-123',
                    'file_name': file_name,
                    'total_revenue': result['total_revenue'],
                    'total_profit': result['total_profit'],
                    'report_url': '/api/report/test-123'
                }

                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps(response, ensure_ascii=False).encode('utf-8'))

                print(f"[SUCCESS] 分析完成: 营业额={result['total_revenue']}, 利润={result['total_profit']}", flush=True)

            except Exception as e:
                print(f"[ERROR] {e}", flush=True)
                import traceback
                traceback.print_exc()
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}, ensure_ascii=False).encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

    def do_GET(self):
        """处理GET请求"""
        if self.path == '/api/records':
            try:
                from database import SessionLocal, AnalysisRecord
                from sqlalchemy import desc
                db = SessionLocal()
                records = db.query(AnalysisRecord).order_by(desc(AnalysisRecord.upload_time)).limit(10).all()
                total = db.query(AnalysisRecord).count()

                data = {
                    'records': [
                        {
                            'id': r.id,
                            'file_name': r.file_name,
                            'upload_time': r.upload_time.strftime("%Y-%m-%d %H:%M"),
                            'total_revenue': r.total_revenue,
                            'total_profit': r.total_profit,
                        }
                        for r in records
                    ],
                    'total': total
                }
                db.close()

                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))
            except Exception as e:
                print(f"[ERROR] {e}", flush=True)
                self.send_response(500)
                self.end_headers()
        elif self.path.startswith('/api/report/'):
            # 简化处理：返回提示信息
            report_id = self.path.replace('/api/report/', '').replace('/download', '')
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_lines()
            html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>报告</title></head>
<body style="padding:40px;font-family:sans-serif;text-align:center">
<h1>📊 分析报告</h1><p>报告ID: {report_id}</p>
<p><a href="/">返回首页</a></p>
</body></html>"""
            self.wfile.write(html.encode('utf-8'))
        else:
            super().do_GET()

    def log_message(self, format, *args):
        print(f"[*] {self.address_string()} - {args[0]}", flush=True)


def main():
    print("="*60)
    print("📊 销售数据分析网站")
    print("="*60)
    print(f"访问地址: http://localhost:{PORT}")
    print("按 Ctrl+C 停止服务\n")

    with socketserver.TCPServer(("", PORT), AnalysisHandler) as httpd:
        print(f"✅ 服务已启动 (端口 {PORT})\n")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n⏹️  已停止")


if __name__ == '__main__':
    main()
