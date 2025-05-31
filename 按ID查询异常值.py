from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import os
import sys

# 获取项目根目录的绝对路径
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(project_root)

# 导入模块
def check_right_table(driver, wait, patient_info):
    """检查右侧表格并返回异常数据"""
    abnormal_records = []
    
    # 等待右侧表格加载
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[col-id='QUALITATIVE_RESULT']")))
    time.sleep(1)  # 额外等待以确保数据加载完成
    
    # 查找所有包含箭头的单元格
    arrow_cells = driver.find_elements(By.CSS_SELECTOR, "div[col-id='QUALITATIVE_RESULT']")
    
    for cell in arrow_cells:
        cell_text = cell.text
        if "↑" in cell_text or "↓" in cell_text or "+" in cell_text or "阳" in cell_text:
            # 获取当前行的其他信息
            row = cell.find_element(By.XPATH, "./ancestor::div[@role='row']")
            cells = row.find_elements(By.CSS_SELECTOR, "div[role='gridcell']")
            
            record = {
                '患者姓名': patient_info['姓名'],
                '患者ID': patient_info['ID'],
                '检验目的': patient_info['检验目的'],
                '报告日期': patient_info['报告日期'],
                '检验项目': cells[1].text if len(cells) > 0 else '',
                '检验结果': cells[2].text if len(cells) > 1 else '',
                '参考范围': cells[4].text if len(cells) > 2 else '',
                '异常标记': cell_text
            }
            abnormal_records.append(record)
    
    return abnormal_records

def read_patient_ids(filename):
    """读取受试者ID文件"""
    with open(filename, 'r', encoding='utf-8') as file:
        return [line.strip() for line in file if line.strip()]

def query_by_ids():
    try:
        # 设置Chrome选项
        chrome_options = Options()
        chrome_options.add_argument('--start-maximized')  # 启动时最大化窗口
        chrome_options.add_argument('--disable-gpu')  # 禁用GPU加速
        chrome_options.add_argument('--no-sandbox')  # 禁用沙箱模式
        chrome_options.add_argument('--disable-dev-shm-usage')  # 禁用/dev/shm使用
        
        # 使用Selenium Manager自动管理驱动程序
        driver = webdriver.Chrome(options=chrome_options)
        print("启动浏览器成功")
        
        # 读取受试者ID
        patient_ids = read_patient_ids("受试者ID.txt")
        print(f"共读取到 {len(patient_ids)} 个受试者ID")
        
        url = "http://192.168.2.66:18110/?APP_ID=CB_HISH01&ENCRY_DATA=39dd4c27a2750549793632944ad6a66c5a6c15d8d23e6ee84d94cfafd8f2073ea42d3c752f5e7753677996e61456a0e019765dc288ae7ad01f8d4b63e6b9e1ff05f52f1138e2a5a29d9b2d0bcb51fe9b2f88a131fde3997fb20adb9a7657fa41fb552dec72813756bcbd092253ec4ce8baad94b809acc37904ad24824b2d67ab"
        
        all_abnormal_records = []
        wait = WebDriverWait(driver, 20)

        # 为每个ID执行查询
        for index, patient_id in enumerate(patient_ids, 1):
            try:
                driver.get(url)
                print(f"\n正在处理第 {index}/{len(patient_ids)} 个ID: {patient_id}")
                
                # 等待输入框可点击
                id_input = wait.until(EC.element_to_be_clickable((By.XPATH, 
                    "/html/body/div[1]/div/div[1]/div[2]/div/div/div[1]/div/div/div[1]/div/form/div/div[4]/div/div/div/div/input")))
                
                # 清除输入框内容并输入
                driver.execute_script("arguments[0].value = '';", id_input)
                time.sleep(0.5)
                id_input.send_keys(patient_id)
                id_input.send_keys(Keys.RETURN)
                time.sleep(2)

                # 等待主表格加载
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.ag-center-cols-container")))
                time.sleep(3)

                # 获取所有行
                rows = driver.find_elements(By.CSS_SELECTOR, "div.ag-center-cols-container > div.ag-row")
                
                for row_index, current_row in enumerate(rows):
                    try:
                        cells = current_row.find_elements(By.CSS_SELECTOR, "div[role='gridcell']")
                        
                        if len(cells) <= 12:
                            print("跳过空行或不完整的行")
                            continue

                        patient_info = {
                            '姓名': cells[3].text,
                            'ID': cells[4].text,
                            '检验目的': cells[6].text,
                            '报告日期': cells[12].text.strip()
                        }

                        if not patient_info['报告日期']:
                            print("跳过空日期记录")
                            continue

                        report_date = patient_info['报告日期'].split()[0]
                        if report_date == '2025-05-06':
                            print(f"处理记录: {patient_info['姓名']}, ID: {patient_info['ID']}")
                            
                            # 点击当前行
                            driver.execute_script("arguments[0].click();", current_row)
                            time.sleep(2)
                            
                            # 检查右侧表格
                            abnormal_records = check_right_table(driver, wait, patient_info)
                            if abnormal_records:
                                all_abnormal_records.extend(abnormal_records)
                                print(f"找到 {len(abnormal_records)} 条异常记录")
                            else:
                                print("未发现异常值")
                        else:
                            print(f"跳过日期为 {patient_info['报告日期']} 的记录")

                    except Exception as e:
                        print(f"处理记录时出错: {str(e)}")
                        continue

            except Exception as e:
                print(f"处理ID {patient_id} 时出错: {str(e)}")
                continue

        # 保存结果
        if all_abnormal_records:
            try:
                df = pd.DataFrame(all_abnormal_records)
                output_file = f'ID查询异常结果_{time.strftime("%Y%m%d_%H%M%S")}.xlsx'
                df.to_excel(output_file, index=False)
                print(f"\n结果已保存到 {output_file}")
                print(f"共找到 {len(all_abnormal_records)} 条异常记录")
            except Exception as e:
                print(f"保存文件时出错: {str(e)}")
        else:
            print("\n未找到任何异常记录")

    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        import traceback
        print(traceback.format_exc())
    
    finally:
        if 'driver' in locals():
            driver.quit()
            print("\n已关闭浏览器")

if __name__ == "__main__":
    query_by_ids()