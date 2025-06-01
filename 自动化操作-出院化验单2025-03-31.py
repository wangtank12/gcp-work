import pyautogui
import time

import win32gui
import win32con
import keyboard
from datetime import datetime
import pygetwindow as gw
import os
# 在文件开头添加
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # 根据实际安装路径修改


# 在文件开头添加新的导入
import tkinter as tk
from tkinter import simpledialog
from paddleocr import PaddleOCR

def read_ids(file_path):
    """读取ID文件"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            # 读取所有行并去除空行和空白字符
            ids = [line.strip() for line in f.readlines() if line.strip()]
        print(f"成功读取 {len(ids)} 个ID")
        return ids
    except Exception as e:
        print(f"读取ID文件失败: {e}")
        return []

def list_all_windows():
    """列出所有窗口标题"""
    try:
        all_windows = gw.getAllWindows()
        print("\n当前所有窗口:")
        for i, window in enumerate(all_windows, 1):
            if window.title:
                # 添加窗口句柄显示
                print(f"{i}. [HWND:{window._hWnd}] {window.title}")

    except Exception as e:
        print(f"获取窗口列表失败: {e}")


def activate_window():
    """激活指定窗口"""
    max_retries = 5
    target_keyword = "门诊医生工作站系统"
    
    for attempt in range(1, max_retries+1):
        try:
            # 获取所有候选窗口
            all_windows = gw.getWindowsWithTitle(target_keyword)
            target_windows = [w for w in all_windows if w.title and target_keyword in w.title]
            
            if not target_windows:
                print(f"第 {attempt} 次尝试：未找到包含'{target_keyword}'的窗口")
                time.sleep(1)
                continue
                
            # 选择第一个可见窗口（更可靠的方式）
            window = None
            for w in target_windows:
                if w.visible:
                    window = w
                    break
            if not window:
                window = target_windows[0]
            
            # 使用win32api强制激活并最大化
            hwnd = window._hWnd
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)  # 改为最大化命令
            win32gui.SetForegroundWindow(hwnd)
            
            # 增加等待时间确保窗口完成最大化
            time.sleep(1)  # 从0.5秒延长到1秒
            
            # 验证激活状态
            time.sleep(0.5)
            active_hwnd = win32gui.GetForegroundWindow()
            if active_hwnd == hwnd:
                print(f"成功激活窗口：{window.title}")
                return True
            else:
                print(f"激活验证失败，当前活动窗口：{win32gui.GetWindowText(active_hwnd)}")
                
        except Exception as e:
            print(f"窗口激活异常[{attempt}]: {str(e)}")
            time.sleep(1)
    
    print("激活失败，当前可用窗口：")
    list_all_windows()
    return False

def create_medical_record(id_number):
    """建立病历"""
    try:
        # 双击选中内容
        pyautogui.click(677, 56)
        time.sleep(1)

        pyautogui.press('delete')
        time.sleep(0.8)
        
        # 输入新ID
        pyautogui.write(id_number)
        time.sleep(0.5)
        
        # 两次回车
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(0.5)
        
        # 挂号相关操作
        pyautogui.click(1130, 347)
        time.sleep(0.5)
        pyautogui.click(1142, 383)
        time.sleep(0.5)
        pyautogui.click(1069, 568)
        time.sleep(0.5)
        pyautogui.click(1019, 597)
        time.sleep(0.5)
        
        # 其他操作
        pyautogui.click(396, 111)
        time.sleep(0.8)
        pyautogui.click(453, 111)
        time.sleep(0.8)
        pyautogui.click(373, 285)
        time.sleep(0.8)
        pyautogui.click(1122, 696)
        time.sleep(0.8)
        
        # 诊断操作
        pyautogui.click(578, 284)
        time.sleep(0.3)
        pyautogui.click(578, 284)  # 双击
        time.sleep(0.5)
        
        # 确认操作
        pyautogui.click(1411, 843)
        time.sleep(0.5)
        pyautogui.click(998, 193)
        time.sleep(0.5)
        pyautogui.press('enter')
        
        print("病历建立完成")
        
    except Exception as e:
        print(f"建立病历时出错: {e}")

def check_quit():
    """检查是否按下q键退出"""
    return keyboard.is_pressed('q')

def create_examination_order():
    """开检查单"""
    try:
        # 点击检查单按钮
        pyautogui.click(633, 105)
        time.sleep(1)
        
        # 选择检查项目
        pyautogui.click(424, 613)
        time.sleep(1)
        
        # 双击某个选项
        pyautogui.click(677, 347)
        time.sleep(0.2)
        pyautogui.click(677, 347)  # 双击
        time.sleep(1)
        
        # 点击GCP试验按钮（双击）
        pyautogui.click(474, 881)
        time.sleep(0.2)
        pyautogui.click(474, 881)  # 双击
        time.sleep(1)
        
        # 点击确定按钮
        pyautogui.press('enter')
        time.sleep(1)
        
        # 点击保存检查
        pyautogui.click(1488, 869)
        time.sleep(1)
        
        # 处理提示框，点击"否"
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('enter')
        
        print("检查单开立完成")
        return True
        
    except Exception as e:
        print(f"开检查单时出错: {e}")
        return False

def check_popup_window():
    """检测新弹出的窗口"""
    try:
        time.sleep(0.5)  # 等待可能的弹窗出现
        windows = gw.getAllWindows()
        for window in windows:
            if window.title and "系统提示" in window.title:
                return True
        return False
    except Exception as e:
        print(f"检查弹窗时出错: {e}")
        return False

# 在文件开头新增配置字典
LAB_ITEMS = {
    "血尿生化": (429, 716),
    "血尿生化电解质": (429, 663),
    "凝血四项": (426, 616),
    "感染凝血": (432, 669),
    "女性妊娠": (406, 715)
}

# 修改项目选择对话框部分（替换原来的循环askyesno）
def main():
    pyautogui.FAILSAFE = True
    
    # 创建隐藏的Tkinter根窗口
    root = tk.Tk()
    root.withdraw()
    
    # 新建项目选择窗口
    select_win = tk.Toplevel(root)
    select_win.title("化验项目选择")
    
    # 创建滚动区域
    canvas = tk.Canvas(select_win)
    scrollbar = tk.Scrollbar(canvas, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)
    
    # 配置画布滚动
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    # 动态生成复选框
    check_vars = {}
    for idx, item in enumerate(LAB_ITEMS):
        var = tk.BooleanVar(value=True)  # 默认选中
        check = tk.Checkbutton(scrollable_frame, text=item, variable=var)
        check.pack(anchor='w', padx=5, pady=2)
        check_vars[item] = var
    
    # 确认按钮
    def on_confirm():
        select_win.destroy()
    
    tk.Button(select_win, text="确认", command=on_confirm).pack(pady=10)
    
    # 窗口居中显示
    select_win.update_idletasks()
    width = 300
    height = min(400, len(LAB_ITEMS)*30 + 100)
    x = (select_win.winfo_screenwidth() - width) // 2
    y = (select_win.winfo_screenheight() - height) // 2
    select_win.geometry(f"{width}x{height}+{x}+{y}")
    
    # 等待用户选择
    select_win.wait_window()
    
    # 收集选中的项目
    selected_items = [item for item, var in check_vars.items() if var.get()]
    
    # 替换原来的 simpledialog.askstring 调用部分
    # 弹出输入对话框获取ID
    input_win = tk.Toplevel(root)
    input_win.title("ID输入")
    
    # 定义 input_ids 变量
    input_ids = ""
    
    # 创建多行文本框
    text_area = tk.Text(input_win, width=40, height=10)
    text_area.pack(padx=10, pady=10)
    
    # 添加粘贴提示
    tip_label = tk.Label(input_win, text="可粘贴多行ID或逗号分隔的ID")
    tip_label.pack()
    
    # 确认按钮
    def on_input_confirm():
        nonlocal input_ids
        input_ids = text_area.get("1.0", "end-1c")
        input_win.destroy()
    
    tk.Button(input_win, text="确认", command=on_input_confirm).pack(pady=5)
    
    # 窗口居中
    input_win.update_idletasks()
    width = 400
    height = 250
    x = (input_win.winfo_screenwidth() - width) // 2
    y = (input_win.winfo_screenheight() - height) // 2
    input_win.geometry(f"{width}x{height}+{x}+{y}")
    
    # 等待输入完成
    input_win.wait_window()
    
    if not input_ids:
        print("没有输入ID，程序退出")
        return
    
    # 解析输入的ID
    ids = [id_str.strip() for id_str in input_ids.replace(',', '\n').splitlines() if id_str.strip()]
    
    if not ids:
        print("没有有效的ID，程序退出")
        return
    
    print(f"\n共读取到 {len(ids)} 个ID")
    print("程序将在3秒后开始处理...")
    print("按 'q' 键可随时停止程序")
    time.sleep(3)
    
    # 处理每个ID时传递配置参数
    for i, id_number in enumerate(ids, 1):
        print(f"[{i}/{len(ids)}] {id_number}")
        process_single_id(id_number, selected_items)  # 添加第二个参数
        time.sleep(2)  # ID之间的间隔
    
    print("\n程序执行完成！")

# 删除文件末尾重复的process_single_id定义，保留以下版本
def process_single_id(id_number, lab_config):
    """处理单个ID的完整流程"""
    try:
        print(f"\n正在处理ID: {id_number}")
        
        # 1. 激活窗口
        activate_window()
        time.sleep(1)
        
        # 2. 建立病历
        create_medical_record(id_number)
        time.sleep(1)
        
        # 3. 开检查单
        create_examination_order()
        time.sleep(1)
        
        # 4. 开化验单
        create_lab_order(lab_config)
        
        print(f"ID {id_number} 的操作执行完成")
        
    except Exception as e:
        print(f"处理过程出现异常: {e}")



def create_lab_order(lab_config):
    """开化验单"""
    try:
        # 点击病人信息按钮
       # pyautogui.click(368, 106)
       # time.sleep(3)
        
        # 初始化PaddleOCR
        ocr = PaddleOCR(
            use_angle_cls=True,
            lang="ch",
            show_log=False,
            use_gpu=False  # 如果没有GPU可以设为False
        )
        
        # 识别性别区域 (1196,133)-(1238,163)
        gender_img = pyautogui.screenshot(region=(1404, 46, 24, 22))
        temp_path = "temp_gender.png"
        gender_img.save(temp_path)
        
        # 使用PaddleOCR识别
        result = ocr.ocr(temp_path, cls=True)
        gender = "男"  # 默认值
        
        if result and result[0]:
            text = "".join([line[1][0] for line in result[0]])
            if "女" in text:
                gender = "女"
        
        os.remove(temp_path)
        
        # 点击化验单按钮
        time.sleep(1)        
        pyautogui.click(748, 107)
        time.sleep(1)
        
        # 双击GCP试验按钮
        pyautogui.click(487, 824)
        time.sleep(0.2)
        pyautogui.click(487, 824)  # 双击
        time.sleep(1)
        
        # 按回车确认
        pyautogui.press('enter')
        time.sleep(1)

       

        # ========== 新增动态执行代码 ========== #
        # 遍历所有选中的化验项目
        for item_name in lab_config:
            # 自动跳过女性妊娠项目（当性别为男时）
            if item_name == "女性妊娠" and gender != "女":
                print(f"跳过女性妊娠项目（性别：{gender}）")
                continue
            
            # 获取项目坐标
            x, y = LAB_ITEMS[item_name]
            
            # 执行项目选择
            pyautogui.click(x, y)
            time.sleep(1)
            
            # 点击添加按钮
            pyautogui.click(1161, 501)
            time.sleep(1)
            print(f"已添加项目：{item_name}")

        # ========== 新增保存操作 ========== #
        # 点击保存按钮
        pyautogui.click(1581, 868)
        time.sleep(1)
        
        # 点击否按钮（保存后的提示）
        pyautogui.press('enter')
        time.sleep(1)

        # 检查是否有弹窗
        if check_popup_window():
            # 如果有系统提示弹窗，按回车确认
            pyautogui.press('enter')
            time.sleep(1)
        else:
            print("未检测到预期的系统提示窗口，请检查操作")
            input("请确认情况后按回车键继续...")  # 添加暂停等待用户确认
            return False
        
        print("化验单开立完成")
        return True
        
    except Exception as e:
        print(f"开化验单时出错: {e}")
        return False





if __name__ == "__main__":
    print("程序启动...")
    print("提示：")
    print("1. 将鼠标移动到屏幕左上角可以紧急停止程序")
    print("2. 按 'q' 键可以正常停止程序")
    main()