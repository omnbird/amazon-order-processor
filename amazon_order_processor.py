import pandas as pd
from bs4 import BeautifulSoup
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import json
import re
import socket
import signal
from contextlib import contextmanager

@contextmanager
def timeout(seconds):
    def signal_handler(signum, frame):
        raise TimeoutError("操作超时")
    signal.signal(signal.SIGALRM, signal_handler)
    signal.alarm(seconds)
    try:
        yield
    finally:
        signal.alarm(0)

class AmazonOrderScraper:
    def __init__(self):
        self.driver = None

    def init_browser(self):
        """初始化浏览器，连接到已打开的Chrome实例"""
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                print(f"尝试连接Chrome浏览器 (第{retry_count + 1}次)...")
                
                # 设置Chrome选项
                options = Options()
                options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
                
                # 使用webdriver_manager自动管理ChromeDriver
                service = Service(ChromeDriverManager().install())
                
                # 设置超时
                with timeout(10):  # 10秒超时
                    self.driver = webdriver.Chrome(service=service, options=options)
                
                print("成功连接到Chrome浏览器")
                return
            except TimeoutError:
                retry_count += 1
                print("连接超时")
                if retry_count < max_retries:
                    print("等待5秒后重试...")
                    time.sleep(5)
                else:
                    print("达到最大重试次数，请检查：")
                    print("1. Chrome是否已使用正确的参数启动")
                    print("2. 端口9222是否被占用")
                    print("3. 是否有其他程序正在使用Chrome")
                    raise Exception("无法连接到Chrome浏览器")
            except Exception as e:
                retry_count += 1
                print(f"连接失败: {str(e)}")
                if retry_count < max_retries:
                    print("等待5秒后重试...")
                    time.sleep(5)
                else:
                    print("达到最大重试次数，请检查：")
                    print("1. Chrome是否已使用正确的参数启动")
                    print("2. 端口9222是否被占用")
                    print("3. 是否有其他程序正在使用Chrome")
                    raise Exception("无法连接到Chrome浏览器")

    def get_orders_page(self, url):
        """获取订单页面内容"""
        try:
            print(f"正在访问URL: {url}")
            self.driver.get(url)
            print("页面加载中...")
            
            # 增加等待时间到30秒
            print("等待页面基本元素加载...")
            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            # 额外等待以确保动态内容加载完成
            print("等待动态内容加载...")
            time.sleep(5)
            
            # 保存页面截图
            print("保存页面截图...")
            self.driver.save_screenshot("page_screenshot.png")
            print("截图已保存为 page_screenshot.png")
            
            # 获取页面内容
            print("获取页面内容...")
            html = self.driver.page_source
            print(f"页面内容长度: {len(html)} 字符")
            return html
            
        except Exception as e:
            print(f"获取页面时出错: {str(e)}")
            print("错误详情:", e.__class__.__name__)
            import traceback
            print("错误堆栈:", traceback.format_exc())
            return None

    def extract_addresses(self, html_content):
        """从HTML内容中提取地址"""
        if not html_content:
            return []
        
        soup = BeautifulSoup(html_content, 'html.parser')
        addresses = []
        
        # 查找所有包含地址信息的div
        address_divs = soup.find_all('div', {'data-test-id': 'shipping-section-buyer-address'})
        
        for div in address_divs:
            spans = div.find_all('span')
            address_parts = []
            for span in spans:
                text = span.get_text(strip=True)
                if text:
                    address_parts.append(text)
            
            if address_parts:
                address_parts = [part for part in address_parts if part]
                full_address = ' '.join(address_parts)
                addresses.append(full_address)
        
        return addresses

    def close(self):
        """关闭浏览器"""
        if self.driver:
            self.driver.quit()

def update_excel_with_addresses(excel_path, addresses):
    """更新Excel文件中的地址信息"""
    try:
        df = pd.read_excel(excel_path)
        
        if '配送地址' not in df.columns:
            df['配送地址'] = ''
        
        for i, address in enumerate(addresses):
            if i < len(df):
                df.at[i, '配送地址'] = address
        
        df.to_excel(excel_path, index=False)
        print(f"成功更新了 {len(addresses)} 个地址到Excel文件")
    except Exception as e:
        print(f"更新Excel文件时出错: {str(e)}")

# 提取城市（只取到市/区/郡结尾）
def extract_city(span1):
    match = re.search(r'([\u4e00-\u9fa5A-Za-z0-9]+?[市区郡])', span1)
    if match:
        city = match.group(1)
    else:
        city = ''
    return city

# 解析单个订单详情页
def extract_order_info_from_html(html):
    soup = BeautifulSoup(html, 'html.parser')
    order_id_tag = soup.find('span', {'data-test-id': 'order-id-value'})
    order_id = order_id_tag.get_text(strip=True) if order_id_tag else ''
    phone_tag = soup.find('span', {'data-test-id': 'shipping-section-phone'})
    phone = phone_tag.get_text(strip=True) if phone_tag else ''
    address_div = soup.find('div', {'data-test-id': 'shipping-section-buyer-address'})
    spans = address_div.find_all('span', recursive=False) if address_div else []  # 只获取直接子span
    if len(spans) >= 5:
        # 固定位置：第一个是姓名，倒数三个分别是省/州、邮编、国家
        name = spans[0].get_text(strip=True)
        # 处理省/州，只取第一个文本节点
        province_span = spans[-3]
        province = next((text for text in province_span.stripped_strings), '')
        zipcode = spans[-2].get_text(strip=True)
        country = spans[-1].get_text(strip=True)
        
        # 中间的spans是详细地址（可能有1-3个span）
        address_lines = [spans[i].get_text(strip=True) for i in range(1, len(spans)-3)]
        address = ' '.join(address_lines)
        
        # 从第一个地址行提取城市（通常包含"市"）
        city_match = re.search(r'([\u4e00-\u9fa5A-Za-z0-9]+?[市区郡])', address_lines[0]) if address_lines else None
        city = city_match.group(1) if city_match else ''
    else:
        name = city = province = zipcode = address = country = ''
    return {
        '订单号': order_id,
        '收件人姓名': name,
        '收件人电话': phone,
        '收件人国家': country,
        '收件人省/州': province,
        '收件人城市': city,
        '收件人邮编': zipcode,
        '收件人地址': address,
        
    }

def check_chrome_running():
    """检查Chrome是否正在运行并监听端口9222"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        result = sock.connect_ex(('127.0.0.1', 9222))
        return result == 0
    except:
        return False
    finally:
        sock.close()

def main():
    print("请先执行以下步骤：")
    print("1. 关闭所有Chrome窗口")
    print("2. 在终端中运行以下命令：")
    print('/Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome --remote-debugging-port=9222 --user-data-dir="/tmp/chrome-debug"')
    print("3. 在新打开的Chrome中登录亚马逊卖家后台")
    print("\n注意：确保Chrome已经启动并且端口9222没有被占用")
    
    if not check_chrome_running():
        print("错误：Chrome未启动或未正确配置")
        print("请确保已使用正确的参数启动Chrome")
        return
    
    input("完成上述步骤后，按回车键继续...")

    print("请输入亚马逊订单详情页URL（每行一个，输入空行结束）：")
    urls = []
    while True:
        line = input()
        if not line.strip():
            break
        urls.append(line.strip())
    
    scraper = AmazonOrderScraper()
    results = []
    try:
        print("正在初始化浏览器...")
        scraper.init_browser()
        print("浏览器初始化完成")
        
        for amazon_url in urls:
            print(f"正在访问订单详情页: {amazon_url}")
            html = scraper.get_orders_page(amazon_url)
            if html:
                print("开始解析订单信息...")
                result = extract_order_info_from_html(html)
                print("解析结果：")
                for k, v in result.items():
                    print(f"{k}: {v}")
                print("-"*40)
                results.append(result)
            else:
                print("未能获取页面内容")
        
        # 可选：写入Excel
        # df = pd.DataFrame(results)
        # df.to_excel('amazon_orders.xlsx', index=False)
        # print('所有订单信息已写入 amazon_orders.xlsx')

        columns = [
            '订单号', '平台交易号', '交接仓', '产品名称', '收件人姓名', '收件人电话', '收件人邮箱', '收件人税号', '收件人公司',
            '收件人国家', '收件人省/州', '收件人城市', '收件人邮编', '收件人地址', '收件人门牌号', '销售平台', '发件人税号信息', 'CSP', '包装尺寸【长】cm', '包装尺寸【宽】cm', '包装尺寸【高】cm', '收款到账日期', '币种类型', '是否含电',
            '拣货单信息', 'IOSS税号', '中文品名1', '英文品名1', '单票数量1', '重量1(g)', '申报价值1'
        ]


        # 组装数据
        data = []
        for item in results:
            row = [
                item.get('订单号', ''),
                '',  # 平台交易号
                '深圳燕文',  # 交接仓
                '燕文专线快递-普货',  # 产品名称
                item.get('收件人姓名', ''),
                item.get('收件人电话', ''),
                '',  # 收件人邮箱
                '',  # 收件人税号
                '',  # 收件人公司
                item.get('收件人国家', ''),
                item.get('收件人省/州', ''),
                item.get('收件人城市', ''),
                item.get('收件人邮编', ''),
                item.get('收件人地址', ''),
                '',  # 收件人门牌号
                '',  # 销售平台
                '',  # 发件人税号信息
                '',  # CSP
                '',  # 包装尺寸【长】cm
                '',  # 包装尺寸【宽】cm
                '',  # 包装尺寸【高】cm
                '',  # 收款到账日期
                '美元',  # 币种类型
                '否',  # 是否含电
                '',  # 拣货单信息
                '',  # IOSS税号
                '舞蹈服',  # 中文品名1
                'dance-suit',  # 英文品名1
                '1',  # 单票数量1
                '300',  # 重量1(g)
                '5',  # 申报价值1
            ]
            data.append(row)

        df = pd.DataFrame(data, columns=columns)
        df.to_excel('output.xlsx', index=False)
        print('已生成 output.xlsx，表头和内容顺序与截图一致。')
                
    except Exception as e:
        print(f"发生错误: {str(e)}")
        print("错误详情:", e.__class__.__name__)
        import traceback
        print("错误堆栈:", traceback.format_exc())
        try:
            scraper.driver.save_screenshot("error_screenshot.png")
            print("已保存错误截图到 error_screenshot.png")
        except:
            print("无法保存错误截图")
    finally:
        print("正在关闭浏览器...")
        scraper.close()
        print("程序执行完成")

if __name__ == "__main__":
    main() 