# my_module.py

import re
import datetime


# 檢查發票號碼格式
def check_invoice_number_format(invoice_number):
    # 定義發票號碼格式的正則表達式
    pattern = r"^[A-Z]{2}\d{8}$"
    
    # 使用 re.match() 函數進行格式匹配
    match = re.match(pattern, invoice_number)
    
    if match:
        return True
    else:
        return False

# 產生傳票編號
def generate_voucher_number():
    now = datetime.datetime.now()
    year = now.year
    month = now.month
    day = now.day
    hour = now.hour
    minute = now.minute
    second = now.second
    microsecond = now.microsecond

    # 生成傳票號碼
    voucher_number = "{:04d}{:02d}{:02d}{:02d}{:02d}{:02d}".format(
        year, month, day, hour, minute, second
    )

    return voucher_number

# 找出指定工作表中 A 欄最後一個非空單元格的位置，並回傳下一個單元格的索引    
def find_last_empty_row(sh):
    max_row = sh.max_row
    for i in range(max_row, 0, -1):
        if sh.cell(row=i, column=1).value is not None:
            max_row = i
            break        
    return max_row