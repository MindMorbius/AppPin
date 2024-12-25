import win32gui
import win32con
import win32process
import win32clipboard
import pyautogui
import time
import hashlib
from datetime import datetime
from PIL import ImageGrab
import os
import cv2
import numpy as np
from paddleocr import PaddleOCR
import win32com.client
import traceback

class WeChatMonitor:
    def __init__(self, contact_name, output_file):
        self.contact_name = contact_name
        self.output_file = output_file
        self.message_cache = set()
        self.last_messages = []
        self.last_message = None
        self.screenshot_dir = "screenshots"
        self.wechat_hwnd = None
        self.last_ocr_text = None  # 添加OCR文本缓存
        self.contact_found = False  # 添加联系人状态标记
        
        # 初始化PaddleOCR
        print("初始化OCR引擎...")
        self.ocr = PaddleOCR(use_angle_cls=True, lang='ch', use_gpu=False)
        
        if not os.path.exists(self.screenshot_dir):
            os.makedirs(self.screenshot_dir)
            
    def get_wechat_window(self):
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                # 更精确的微信窗口匹配
                if window_text == "微信" or window_text.lower() == "wechat":
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    self.wechat_pid = pid
                    hwnds.append(hwnd)
                return True
        
        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        # 打印所有找到的窗口信息，便于调试
        for hwnd in hwnds:
            print(f"找到微信窗口: {win32gui.GetWindowText(hwnd)}, Handle: {hwnd}")
        return hwnds[0] if hwnds else None

    def is_window_foreground(self, hwnd):
        try:
            foreground_hwnd = win32gui.GetForegroundWindow()
            return hwnd == foreground_hwnd
        except:
            return False

    def capture_window(self, hwnd):
        # 获取窗口位置
        x, y, x1, y1 = win32gui.GetWindowRect(hwnd)
        screenshot = ImageGrab.grab((x, y, x1, y1))
        return np.array(screenshot)

    def find_contact(self, img):
        debug_path = f"{self.screenshot_dir}/debug_contact.png"
        cv2.imwrite(debug_path, cv2.cvtColor(img, cv2.COLOR_RGB2BGR))
        print(f"已保存调试截图: {debug_path}")
        
        results = self.ocr.ocr(img, cls=True)
        if results:
            print("\n识别到的文本:")
            for line in results[0]:
                text = line[1][0]  # 获取文本
                confidence = line[1][1]
                print(f"文本: {text}, 置信度: {confidence:.2f}")
                
                # 将下划线替换为空格进行比较
                if self.contact_name.replace('_', ' ') in text.replace('_', ' '):
                    box = line[0]
                    center_x = sum(p[0] for p in box) / 4
                    center_y = sum(p[1] for p in box) / 4
                    print(f"找到联系人 {text} 在位置: ({center_x}, {center_y})")
                    
                    debug_img = cv2.imread(debug_path)
                    cv2.circle(debug_img, (int(center_x), int(center_y)), 5, (0, 0, 255), -1)
                    cv2.polylines(debug_img, [np.array(box, np.int32)], True, (0, 255, 0), 2)
                    cv2.imwrite(f"{self.screenshot_dir}/debug_contact_marked.png", debug_img)
                    
                    return (center_x, center_y)
        
        print(f"未找到联系人: {self.contact_name}")
        return None

    def find_new_message(self, img):
        # 保存HSV调试图像
        hsv = cv2.cvtColor(img, cv2.COLOR_RGB2HSV)
        cv2.imwrite(f"{self.screenshot_dir}/debug_hsv.png", hsv)
        
        # 白色气泡的HSV范围
        lower_white = np.array([0, 0, 200])
        upper_white = np.array([180, 30, 255])
        
        # 创建掩码并保存
        mask = cv2.inRange(hsv, lower_white, upper_white)
        cv2.imwrite(f"{self.screenshot_dir}/debug_mask.png", mask)
        
        # 形态学操作
        kernel = np.ones((5,5), np.uint8)
        mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)
        cv2.imwrite(f"{self.screenshot_dir}/debug_mask_morphology.png", mask)
        
        # 查找轮廓
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        # 在原图上标记轮廓
        debug_img = img.copy()
        cv2.drawContours(debug_img, contours, -1, (0, 255, 0), 2)
        cv2.imwrite(f"{self.screenshot_dir}/debug_contours.png", cv2.cvtColor(debug_img, cv2.COLOR_RGB2BGR))
        
        # 筛选合适大小的轮廓，且必须在左侧区域
        valid_contours = []
        img_width = img.shape[1]
        for cnt in contours:
            x, y, w, h = cv2.boundingRect(cnt)
            # 只选择左半部分的消息
            if w > 50 and h > 20 and x < img_width/2:
                valid_contours.append((y, x, w, h))
                print(f"找到有效轮廓: x={x}, y={y}, w={w}, h={h}")
        
        if valid_contours:
            # 按y坐标排序，获取最新的消息（最下方的）
            valid_contours.sort(reverse=True)
            y, x, w, h = valid_contours[0]
            
            # 调整识别区域，向右扩展以包含完整消息
            x_right = min(x + w + 100, img_width)  # 向右扩展100像素
            center_x = x + w//2
            center_y = y + h//2
            
            # 标记选中的轮廓
            cv2.rectangle(debug_img, (x, y), (x_right, y+h), (0, 0, 255), 2)
            cv2.circle(debug_img, (center_x, center_y), 5, (255, 0, 0), -1)
            cv2.imwrite(f"{self.screenshot_dir}/debug_selected_contour.png", cv2.cvtColor(debug_img, cv2.COLOR_RGB2BGR))
            
            print(f"选择最新消息位置: ({center_x}, {center_y})")
            return (center_x, center_y, x, y, x_right-x, h)  # 返回中心点和区域信息
        
        print("未找到有效的消息气泡")
        return None

    def copy_message(self, pos, img_shape):
        try:
            # 保存当前鼠标位置
            original_x, original_y = pyautogui.position()
            
            # 获取窗口位置
            wx, wy, _, _ = win32gui.GetWindowRect(self.wechat_hwnd)
            
            # 计算实际点击位置
            click_x = wx + pos[0]
            click_y = wy + pos[1]
            
            # 移动并点击
            pyautogui.moveTo(click_x, click_y)
            time.sleep(0.1)
            pyautogui.click(button='right')
            time.sleep(0.2)
            pyautogui.moveRel(10, 20)
            pyautogui.click()
            time.sleep(0.1)
            
            # 恢复鼠标位置
            pyautogui.moveTo(original_x, original_y)
            
            # 获取剪贴板内容
            win32clipboard.OpenClipboard()
            try:
                text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                return text
            finally:
                win32clipboard.CloseClipboard()
        except Exception as e:
            print(f"复制消息失败: {e}")
            return None

    def save_message(self, message, img):
        if not message:
            return False
            
        msg_hash = hashlib.md5(message.encode()).hexdigest()
        if msg_hash in self.message_cache:
            return False
            
        self.message_cache.add(msg_hash)
        
        # 保存截图
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = f"{self.screenshot_dir}/wechat_{timestamp}.png"
        cv2.imwrite(screenshot_path, cv2.cvtColor(img, cv2.COLOR_RGB2BGR))
        
        # 保存消息和时间
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.output_file, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] 收到新消息:\n{message}\n")
            f.write(f"Screenshot: {screenshot_path}\n")
            f.write("-" * 50 + "\n")
        
        print(f"[{timestamp}] 收到新消息，长度: {len(message)}")
        return True

    def bring_to_front(self):
        if not self.wechat_hwnd:
            return False
            
        try:
            # 强制激活窗口
            win32gui.ShowWindow(self.wechat_hwnd, win32con.SW_RESTORE)
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(self.wechat_hwnd)
            
            # 验证是否成功置前
            time.sleep(0.5)  # 等待窗口状态更新
            return self.is_window_foreground(self.wechat_hwnd)
        except Exception as e:
            print(f"置前失败: {e}")
            return False

    def find_chat_area(self, img):
        """识别对话区域和联系人列表区域，使用遮罩而不是裁切"""
        height, width = img.shape[:2]
        # 一般情况下，联系人列表在左侧约1/4位置
        contact_list_width = width // 4
        
        # 创建遮罩
        mask = np.zeros(img.shape, dtype=np.uint8)
        # 只保留对话区域
        mask[:, contact_list_width:] = img[:, contact_list_width:]
        
        # 保存调试图像
        debug_img = img.copy()
        cv2.line(debug_img, (contact_list_width, 0), (contact_list_width, height), (0, 255, 0), 2)
        cv2.imwrite(f"{self.screenshot_dir}/debug_areas.png", cv2.cvtColor(debug_img, cv2.COLOR_RGB2BGR))
        
        return mask, contact_list_width

    def find_chat_messages(self, img, chat_area_start):
        """直接OCR识别对话区域的���息"""
        # 只处理对话区域
        chat_area = img[:, chat_area_start:]
        
        # 保存调试图像
        cv2.imwrite(f"{self.screenshot_dir}/debug_chat_area.png", 
                    cv2.cvtColor(chat_area, cv2.COLOR_RGB2BGR))
        
        messages = []
        results = self.ocr.ocr(chat_area, cls=True)
        
        if results and results[0]:
            print("\n识别到的文本:")
            for line in results[0]:
                box = line[0]  # 文本框坐标
                text = line[1][0]  # 文本内容
                confidence = line[1][1]  # 置信度
                
                # 计算文本框中心点的x坐标（相对于对话区域）
                center_x = sum(p[0] for p in box) / 4
                
                # 判断是否为对方消息（在左半部分）
                is_received = center_x < chat_area.shape[1] / 2
                
                # 转换回原图坐标
                abs_center_x = center_x + chat_area_start
                abs_center_y = sum(p[1] for p in box) / 4
                
                print(f"文本: {text}")
                print(f"位置: {abs_center_x:.0f}, {abs_center_y:.0f}")
                print(f"类型: {'对方' if is_received else '自己'}")
                
                messages.append({
                    'type': 'received' if is_received else 'sent',
                    'text': text,
                    'box': box,
                    'center': (abs_center_x, abs_center_y),
                    'confidence': confidence
                })
        
        # 按y坐标排序，获取最新消息
        messages.sort(key=lambda x: x['center'][1], reverse=True)
        return messages

    def check_current_contact(self, img):
        """检查当前对话框是否为目标联系人"""
        height, width = img.shape[:2]
        # 检查对话框顶部区域（通常是联系人名称位置）
        top_area = img[0:50, width//4:width]
        
        # 保存调试图像
        cv2.imwrite(f"{self.screenshot_dir}/debug_top_area.png", cv2.cvtColor(top_area, cv2.COLOR_RGB2BGR))
        
        results = self.ocr.ocr(top_area, cls=True)
        if results and results[0]:
            for line in results[0]:
                text = line[1][0]
                print(f"当前对话框联系人: {text}")
                if self.contact_name.replace('_', ' ') in text.replace('_', ' '):
                    return True
        return False

    def find_contact_in_list(self, img):
        """在联系人列表中查找并点击目标联系人"""
        height, width = img.shape[:2]
        contact_list_area = img[:, :width//4]
        
        # 保存调试图像
        cv2.imwrite(f"{self.screenshot_dir}/debug_contact_list.png", cv2.cvtColor(contact_list_area, cv2.COLOR_RGB2BGR))
        
        results = self.ocr.ocr(contact_list_area, cls=True)
        if results and results[0]:
            for line in results[0]:
                text = line[1][0]
                if self.contact_name.replace('_', ' ') in text.replace('_', ' '):
                    box = line[0]
                    center_x = sum(p[0] for p in box) / 4
                    center_y = sum(p[1] for p in box) / 4
                    print(f"找到联系人 {text} 在位置: ({center_x}, {center_y})")
                    
                    # 点击联系人
                    wx, wy, _, _ = win32gui.GetWindowRect(self.wechat_hwnd)
                    pyautogui.click(wx + center_x, wy + center_y)
                    time.sleep(0.5)  # 等待对话框加载
                    return True
        return False

    def find_new_messages(self, current_messages):
        """比较并找出新消息"""
        # 提取当前所有对方消息的文本
        current_texts = {
            msg['text'] for msg in current_messages 
            if msg['type'] == 'received' and not msg['text'].startswith('"') and not ':' in msg['text']
        }
        
        # 找出新消息（在当前消息中但不在缓存中的）
        new_messages = current_texts - self.message_cache
        
        # 更新缓存
        self.message_cache.update(current_texts)
        
        return new_messages

    def run(self):
        print(f"开始监控联系人 {self.contact_name} 的消息...")
        last_check_time = time.time()
        check_interval = 5
        
        while True:
            try:
                current_time = time.time()
                if current_time - last_check_time < check_interval:
                    time.sleep(0.5)
                    continue
                    
                last_check_time = current_time
                
                if not self.wechat_hwnd:
                    self.wechat_hwnd = self.get_wechat_window()
                    if not self.wechat_hwnd:
                        continue
                
                if not self.bring_to_front():
                    continue
                
                # 捕获完整窗口内容
                img = self.capture_window(self.wechat_hwnd)
                
                # 只在未找到联系人时进行检查
                if not self.contact_found:
                    if not self.check_current_contact(img):
                        print("当前不是目标联系人对话框，尝试在列表中查找...")
                        if not self.find_contact_in_list(img):
                            print(f"未找到联系人: {self.contact_name}")
                            time.sleep(2)
                            continue
                        img = self.capture_window(self.wechat_hwnd)
                    self.contact_found = True
                    print(f"已找到联系人 {self.contact_name}，开始监控消息...")
                
                # 获取对话区域起始位置
                _, chat_start_x = self.find_chat_area(img)
                
                # 直接OCR识别消息
                current_messages = self.find_chat_messages(img, chat_start_x)
                
                # 检查新消息
                new_messages = self.find_new_messages(current_messages)
                
                if new_messages:
                    print(f"检测到 {len(new_messages)} 条新消息:")
                    for text in new_messages:
                        print(f"新消息: {text}")
                        # 找到对应的消息对象以获取位置信息
                        for msg in current_messages:
                            if msg['text'] == text and msg['type'] == 'received':
                                # 尝试复制获取完整文本
                                full_text = self.copy_message(msg['center'], img.shape)
                                if full_text:
                                    self.save_message(full_text, img)
                                break
                else:
                    print("未检测到新消息")
                
                print(f"等待 {check_interval} 秒后进行下一次检查...")
                
            except KeyboardInterrupt:
                print("\n监控已停止")
                break
            except Exception as e:
                print(f"发生错误: {e}")
                traceback.print_exc()
                time.sleep(2)
                self.contact_found = False

if __name__ == "__main__":
    monitor = WeChatMonitor(
        contact_name="文件传输助手",
        output_file="messages.txt"
    )
    monitor.run()