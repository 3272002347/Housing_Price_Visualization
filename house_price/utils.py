import re

def extract_number(text):
    """从字符串中提取数字（支持小数）"""
    if not text:
        return None
    match = re.search(r'(\d+\.?\d*)', str(text))
    if match:
        return float(match.group(1))
    return None

def clean_area(area_str):
    """清洗面积：提取数字，默认单位平米"""
    num = extract_number(area_str)
    return num if num else None

def clean_house_age(age_str):
    """清洗房龄：提取数字，默认单位年"""
    num = extract_number(age_str)
    return int(num) if num else None

def clean_total_price(price_str):
    """清洗总价：提取数字，默认单位万"""
    num = extract_number(price_str)
    return num if num else None