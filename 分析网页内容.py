from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os

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
        if "↑" in cell_text or "↓" in cell_text or "+" in cell_text:
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

def download_table_data():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    chromedriver_path = os.path.join(current_dir, "chromedriver.exe")
    
    # 读取受试者ID
    patient_ids = read_patient_ids("受试者ID1.txt")
    print(f"共读取到 {len(patient_ids)} 个受试者ID")
    
    try:
        service = Service(chromedriver_path)
        driver = webdriver.Chrome(service=service)
        print("启动浏览器")
        
        driver.maximize_window()
        print("窗口已最大化")
        
        all_abnormal_records = []
        wait = WebDriverWait(driver, 20)
        all_abnormal_records = []
        
        # 等待左侧表格加载
        print("\n等待表格加载...")
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.ag-center-cols-container")))
        
        try:
            # 等待主表格加载完成
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.ag-center-cols-container")))
            time.sleep(3)  # 给更多时间确保表格完全加载
            
            # 获取所有行
            rows = driver.find_elements(By.CSS_SELECTOR, "div.ag-center-cols-container > div.ag-row[role='row']:not(.ag-row-loading)")
            
            # 遍历每一行
            for row_index, current_row in enumerate(rows):
                try:
                    cells = current_row.find_elements(By.CSS_SELECTOR, "div[role='gridcell']")
                    
                    # 直接获取患者信息
                    patient_info = {
                        '姓名': cells[3].text,
                        'ID': cells[4].text,
                        '检验目的': cells[6].text,
                        '报告日期': cells[12].text
                    }
                    
                    print(f"\n处理第 {row_index + 1} 行:")
                    print(f"患者: {patient_info['姓名']}, ID: {patient_info['ID']}")
                    
                    # 点击当前行
                    driver.execute_script("arguments[0].click();", current_row)
                    time.sleep(2)  # 等待右侧表格更新
                    
                    # 检查右侧表格
                    abnormal_records = check_right_table(driver, wait, patient_info)
                    if abnormal_records:
                        all_abnormal_records.extend(abnormal_records)
                        print(f"找到 {len(abnormal_records)} 条异常记录")
                    else:
                        print("未发现异常值")
                
                except Exception as e:
                    print(f"处理第 {row_index + 1} 行时出错: {str(e)}")
                    continue
            
        except Exception as e:
            print(f"处理表格时出错: {str(e)}")
        
        # 保存所有异常记录
        if all_abnormal_records:
            try:
                df = pd.DataFrame(all_abnormal_records)
                output_file = '检验异常结果.xlsx'
                
                try:
                    # 检查文件是否存在且被占用
                    if os.path.exists(output_file):
                        print(f"\n警告：文件 {output_file} 已存在")
                        print("请关闭该文件后重试，或者指定新的文件名")
                        new_output_file = f'检验异常结果_{time.strftime("%Y%m%d_%H%M%S")}.xlsx'
                        print(f"自动保存为新文件：{new_output_file}")
                        df.to_excel(new_output_file, index=False)
                        print(f"\n结果已保存到 {new_output_file}")
                    else:
                        df.to_excel(output_file, index=False)
                        print(f"\n结果已保存到 {output_file}")
                    print(f"共找到 {len(all_abnormal_records)} 条异常记录")
                except PermissionError:
                    # 如果仍然无法保存，尝试保存到用户的文档目录
                    user_docs = os.path.expanduser('~\\Documents')
                    backup_file = os.path.join(user_docs, f'检验异常结果_{time.strftime("%Y%m%d_%H%M%S")}.xlsx')
                    print(f"\n无法保存到当前目录，尝试保存到：{backup_file}")
                    df.to_excel(backup_file, index=False)
                    print(f"结果已保存到 {backup_file}")
                    
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
    download_table_data()