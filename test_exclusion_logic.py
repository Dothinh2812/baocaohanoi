#!/usr/bin/env python3
"""
Script test để kiểm tra logic giảm trừ sau khi sửa
"""

import sys
import os

# Thêm thư mục hiện tại vào path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from exclusion_process import process_exclusion_reports

if __name__ == "__main__":
    print("="*80)
    print("🧪 BẮT ĐẦU TEST LOGIC GIẢM TRỪ SAU KHI SỬA")
    print("="*80)

    # Chạy quy trình giảm trừ
    result = process_exclusion_reports()

    if result:
        print("\n" + "="*80)
        print("✅ QUY TRÌNH GIẢM TRỪ HOÀN THÀNH")
        print("="*80)
        print("\n📊 Kiểm tra kết quả tại:")
        print("   - downloads/kq_sau_giam_tru/So_sanh_C11_SM4.xlsx")
        print("   - downloads/kq_sau_giam_tru/So_sanh_C11_SM2.xlsx")
        print("   - downloads/kq_sau_giam_tru/So_sanh_C12_SM1.xlsx")
        print("   - downloads/kq_sau_giam_tru/So_sanh_C14.xlsx")
        print("   - downloads/kq_sau_giam_tru/So_sanh_C15.xlsx")
    else:
        print("\n" + "="*80)
        print("❌ QUY TRÌNH GIẢM TRỪ GẶP LỖI")
        print("="*80)
        sys.exit(1)
