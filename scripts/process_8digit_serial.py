import pandas as pd
import os
import sys
import re

def extract_8digit_serial_numbers(excel_path):
    """
    从特定格式的Excel文件中提取8位序列号
    只读取第一个工作表
    
    参数:
    excel_path: Excel文件路径
    
    返回:
    序列号列表
    """
    try:
        # 只读取第一个工作表 (八位序列号收集收集结果yd5)
        df = pd.read_excel(excel_path, sheet_name=0)
        
        serial_numbers = set()
        
        # 处理第一个工作表
        # 列结构: A-提交者, B-提交时间, C-序列号
        if '此处填写（必填）' in df.columns:
            serials = df['此处填写（必填）'].dropna().astype(str).str.strip()
            for serial in serials:
                # 提取序列号，忽略点号
                clean_serial = extract_serial(serial)
                if clean_serial:
                    serial_numbers.add(clean_serial)
            print(f"从表1中提取了 {len(serial_numbers)} 个有效序列号")
        else:
            print("错误: 未找到'此处填写（必填）'列")
            # 尝试查找其他可能的列名
            print(f"可用的列: {list(df.columns)}")
            sys.exit(1)
        
        return list(serial_numbers)
        
    except Exception as e:
        print(f"读取Excel文件时出错: {str(e)}")
        sys.exit(1)

def extract_serial(serial_str):
    """
    从字符串中提取序列号，忽略点号
    
    参数:
    serial_str: 可能包含序列号的字符串
    
    返回:
    提取出的序列号，或None
    """
    # 移除所有点号
    clean_str = serial_str.replace('.', '')
    
    # 检查是否只包含十六进制字符
    if re.match(r'^[a-fA-F0-9]+$', clean_str):
        return clean_str.lower()  # 统一转为小写
    
    # 如果包含非十六进制字符，尝试提取十六进制部分
    matches = re.findall(r'[a-fA-F0-9]+', serial_str)
    if matches:
        # 合并所有十六进制部分
        hex_part = ''.join(matches).lower()
        return hex_part
    
    return None

def update_whitelist(serial_numbers, whitelist_path='WhiteList.config'):
    """
    更新白名单文件
    
    参数:
    serial_numbers: 新的序列号列表
    whitelist_path: 白名单文件路径
    """
    try:
        # 读取现有白名单（如果存在）
        existing_serials = set()
        if os.path.exists(whitelist_path):
            with open(whitelist_path, 'r') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#'):  # 忽略空行和注释
                        existing_serials.add(line)
        
        # 添加新的序列号
        new_serials = set(serial_numbers)
        all_serials = existing_serials.union(new_serials)
        
        # 按字母顺序排序
        sorted_serials = sorted(all_serials)
        
        # 写入白名单文件
        with open(whitelist_path, 'w') as f:
            f.write("# 设备白名单列表\n")
            f.write("# 每行一个十六进制设备序列号\n")
            f.write("# 自动从Excel文件表1更新\n\n")
            for serial in sorted_serials:
                f.write(f"{serial}\n")
        
        new_count = len(new_serials - existing_serials)
        print(f"白名单已更新: 总数 {len(sorted_serials)}, 新增 {new_count} 个序列号")
        
        return new_count > 0  # 返回是否有更新
        
    except Exception as e:
        print(f"更新白名单文件时出错: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    # 文件路径配置
    EXCEL_PATH = 'data/八位序列号收集（收集结果）.xlsx'
    
    # 创建数据目录（如果不存在）
    os.makedirs(os.path.dirname(EXCEL_PATH), exist_ok=True)
    
    # 检查Excel文件是否存在
    if not os.path.exists(EXCEL_PATH):
        print(f"警告: Excel文件 {EXCEL_PATH} 不存在")
        sys.exit(0)
    
    # 提取序列号并更新白名单
    serial_numbers = extract_8digit_serial_numbers(EXCEL_PATH)
    
    if not serial_numbers:
        print("未提取到有效的序列号")
        sys.exit(0)
    
    has_changes = update_whitelist(serial_numbers)
    
    # 如果没有变化，退出码为0，有变化则为1
    sys.exit(0 if not has_changes else 1)
