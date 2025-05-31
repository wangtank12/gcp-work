import pyautogui
import time
from PIL import Image
from paddleocr import PaddleOCR
import win32gui
import win32con
import os
from datetime import datetime
import PyPDF2
import win32print
import win32api
import threading
import queue
import shutil

def print_pdf(pdf_path):
    """打印PDF文件"""
    try:
        
        print(f"开始打印文件: {pdf_path}")
        win32api.ShellExecute(
            0,
            "print",
            pdf_path,
            None,
            ".",
            0
        )
        time.sleep(2)
        print("打印命令已发送")
    except Exception as e:
        print(f"打印文件时出错: {e}")
if __name__ == "__main__":
    test_pdf_path = r"D:\工作\阿莫西林\300350841_合并.pdf"  # 替换为实际的PDF文件路径
    print_pdf(test_pdf_path)