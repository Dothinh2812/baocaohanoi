#!/usr/bin/env python
"""
Script để phân tích cấu trúc DOM của trang báo cáo C1.4 và C1.5
"""

from c1_report_download import login_baocao_hanoi
import time
import os

def analyze_page_structure(page_baocao, url, report_name):
    """Phân tích cấu trúc DOM của trang báo cáo"""
    
    print(f"\n{'='*80}")
    print(f"PHÂN TÍCH CẤU TRÚC TRANG: {report_name}")
    print(f"{'='*80}")
    
    # Truy cập URL
    print(f"\n✓ Đang truy cập: {url}")
    page_baocao.goto(url, wait_until="networkidle", timeout=60000)
    time.sleep(5)  # Đợi thêm để trang load hoàn toàn
    print("✅ Đã tải trang thành công")
    
    # Lưu screenshot
    screenshot_dir = "downloads/analysis"
    os.makedirs(screenshot_dir, exist_ok=True)
    screenshot_path = os.path.join(screenshot_dir, f"{report_name}_screenshot.png")
    page_baocao.screenshot(path=screenshot_path, full_page=True)
    print(f"✅ Đã lưu screenshot: {screenshot_path}")
    
    # Phân tích các dropdown buttons
    print("\n--- TÌM KIẾM DROPDOWN BUTTONS ---")
    dropdown_buttons = page_baocao.locator("button").all()
    print(f"Số lượng buttons: {len(dropdown_buttons)}")
    for i, btn in enumerate(dropdown_buttons[:10]):  # Chỉ hiển thị 10 buttons đầu
        try:
            text = btn.inner_text()[:50] if btn.inner_text() else "(no text)"
            print(f"  Button {i}: {text}")
        except:
            pass
    
    # Phân tích các select elements
    print("\n--- TÌM KIẾM SELECT ELEMENTS ---")
    select_elements = page_baocao.locator("select").all()
    print(f"Số lượng select elements: {len(select_elements)}")
    for i, sel in enumerate(select_elements):
        try:
            # Lấy options của select
            options = sel.locator("option").all()
            option_texts = [opt.inner_text()[:30] for opt in options[:5]]
            print(f"  Select {i}: {len(options)} options - {option_texts}")
        except Exception as e:
            print(f"  Select {i}: Lỗi - {e}")
    
    # Tìm div.search-criteria
    print("\n--- TÌM KIẾM DIV.SEARCH-CRITERIA ---")
    search_criteria = page_baocao.locator("div.search-criteria")
    if search_criteria.count() > 0:
        print(f"Tìm thấy {search_criteria.count()} div.search-criteria")
        # Tìm các div con
        child_divs = search_criteria.first.locator("> div").all()
        print(f"Số div con trực tiếp: {len(child_divs)}")
        for i, div in enumerate(child_divs[:10]):
            try:
                # Kiểm tra xem có select trong div này không
                selects = div.locator("select").count()
                buttons = div.locator("button").count()
                inputs = div.locator("input").count()
                text = div.inner_text()[:80].replace('\n', ' ') if div.inner_text() else "(empty)"
                print(f"  div[{i+1}]: selects={selects}, buttons={buttons}, inputs={inputs}, text: {text}")
            except Exception as e:
                print(f"  div[{i+1}]: Error - {e}")
    else:
        print("Không tìm thấy div.search-criteria")
        
    # Tìm ngx-dropdown-treeview
    print("\n--- TÌM KIẾM NGX-DROPDOWN-TREEVIEW ---")
    treeview = page_baocao.locator("ngx-dropdown-treeview")
    print(f"Số lượng ngx-dropdown-treeview: {treeview.count()}")
    
    # Tìm các element có chứa text "Tháng"
    print("\n--- TÌM KIẾM ELEMENT CHỨA TEXT 'THÁNG' ---")
    try:
        thang_elements = page_baocao.get_by_text("Tháng", exact=False).all()
        print(f"Số lượng elements chứa 'Tháng': {len(thang_elements)}")
        for i, el in enumerate(thang_elements[:10]):
            try:
                tag = el.evaluate("el => el.tagName")
                text = el.inner_text()[:50] if el.inner_text() else "(empty)"
                print(f"  Element {i}: <{tag}> - {text}")
            except:
                pass
    except Exception as e:
        print(f"Lỗi khi tìm 'Tháng': {e}")
    
    # Lưu HTML của trang
    html_path = os.path.join(screenshot_dir, f"{report_name}_page.html")
    html_content = page_baocao.content()
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"\n✅ Đã lưu HTML: {html_path}")

def main():
    result = login_baocao_hanoi()
    
    if result:
        page_baocao, browser_baocao, playwright_baocao = result
        
        # Phân tích C1.4
        analyze_page_structure(
            page_baocao, 
            "https://baocao.hanoi.vnpt.vn/report/report-info?id=264107&menu_id=275688",
            "C14"
        )
        
        # Phân tích C1.5
        analyze_page_structure(
            page_baocao, 
            "https://baocao.hanoi.vnpt.vn/report/report-info?id=258310&menu_id=305918",
            "C15"
        )
        
        # Đóng browser
        print("\n=== Đóng browser ===")
        browser_baocao.close()
        playwright_baocao.stop()
    else:
        print("❌ Không thể đăng nhập")

if __name__ == "__main__":
    main()
