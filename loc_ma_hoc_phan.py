"""
Script lọc mã học phần ET và EE từ file Excel TKB
"""

import pandas as pd
import sys
import re


def loc_ma_hoc_phan(file_path):
    """
    Lọc các mã học phần bắt đầu bằng ET hoặc EE
    """
    try:
        # Đọc tất cả sheets từ Excel
        print("📖 Đang đọc file Excel...")
        excel_file = pd.ExcelFile(file_path)
        
        print(f"✅ Tìm thấy {len(excel_file.sheet_names)} sheet(s)")
        
        # Lưu tất cả kết quả (dòng chứa ET/EE) và tất cả mã tìm được
        all_rows = []
        all_codes = set()
        
        # Duyệt qua từng sheet
        for sheet_name in excel_file.sheet_names:
            print(f"\n🔍 Đang xử lý sheet: '{sheet_name}'...")
            
            # Đọc sheet (không dùng header để tránh dòng tiêu đề dài)
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str)
            
            # In thử vài dòng đầu để xem cấu trúc
            if len(df) > 0:
                print(f"   📊 Sheet có {len(df)} dòng, {len(df.columns)} cột")
                # Regex bắt mã HP: bắt đầu bằng ET/EE, theo sau là chữ số/chữ in/ dấu gạch
                code_pattern = re.compile(r"\b(ET|EE)[A-Z0-9-]+\b", re.IGNORECASE)

                # Tạo mask dòng nào có chứa mã ET/EE ở bất kỳ cột nào
                row_has_code = df.apply(
                    lambda row: any(bool(code_pattern.search(str(val))) for val in row.values), axis=1
                )

                matched_rows = df[row_has_code].copy()

                if len(matched_rows) > 0:
                    print(f"   ✅ Tìm thấy {len(matched_rows)} dòng chứa mã ET/EE")

                    # Trích xuất mã từ toàn bộ sheet để tổng hợp danh sách mã duy nhất
                    for val in df.astype(str).values.flatten():
                        for m in code_pattern.findall(str(val)):
                            all_codes.add(m.upper())

                    # Gắn tên sheet để truy vết
                    matched_rows['Sheet'] = sheet_name
                    all_rows.append(matched_rows)
                else:
                    print("   ⚠️ Không tìm thấy mã ET/EE trong sheet này")
        
        # Gộp tất cả dòng khớp
        if all_rows:
            result_df = pd.concat(all_rows, ignore_index=True)
            
            # Lưu ra file Excel (bản gốc, giữ nguyên cấu trúc ô)
            output_file = 'Ma_hoc_phan_ET_EE.xlsx'
            result_df.to_excel(output_file, index=False, header=True)
            print(f"\n✅ Đã lọc được {len(result_df)} dòng")
            print(f"📁 Đã lưu vào file: {output_file}")

            # Tạo thêm bản có header chuẩn theo TKB gốc
            HEADERS = [
                'Kỳ','Trường_Viện_Khoa','Mã_lớp','Mã_lớp_kèm','Mã_HP','Tên_HP','Tên_HP_Tiếng_Anh',
                'Khối_lượng','Ghi_chú','Buổi_số','Thứ','Thời_gian','BĐ','KT','Kíp','Tuần','Phòng',
                'Cần_TN','SLĐK','SL_Max','Trạng_thái','Loại_lớp','Đợt_mở','Mã_QL','Hệ','TeachingType',
                'mainclass','Sessionid','Statusid','Khóa'
            ]

            fixed_df = result_df.copy()
            # Cân bằng số cột theo HEADER: cắt bớt hoặc thêm cột trống
            if fixed_df.shape[1] < len(HEADERS):
                for i in range(len(HEADERS) - fixed_df.shape[1]):
                    fixed_df[f'_extra_{i}'] = ''
            elif fixed_df.shape[1] > len(HEADERS):
                fixed_df = fixed_df.iloc[:, :len(HEADERS)]
            fixed_df.columns = HEADERS

            fixed_file = 'Ma_hoc_phan_ET_EE_fixed.xlsx'
            # Ghi đảm bảo có hàng tiêu đề
            fixed_df.to_excel(fixed_file, index=False, header=True)
            print(f"📁 Đồng thời tạo: {fixed_file} (đã chèn hàng tiêu đề cột)")

            # In danh sách các mã học phần duy nhất (từ regex)
            unique_codes = sorted(all_codes)
            print(f"\n📋 Danh sách mã học phần ET/EE ({len(unique_codes)} mã):")
            for code in unique_codes:
                print(f"   - {code}")
            
            # Lưu danh sách mã vào file text
            with open('Danh_sach_ma_ET_EE.txt', 'w', encoding='utf-8') as f:
                f.write("Danh sách mã học phần ET và EE\n")
                f.write("=" * 50 + "\n\n")
                for code in unique_codes:
                    f.write(f"{code}\n")
            print(f"📝 Đã lưu danh sách mã vào: Danh_sach_ma_ET_EE.txt")
            
            return fixed_df
        else:
            print("\n⚠️ Không tìm thấy mã học phần ET hoặc EE nào!")
            return None
            
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        import traceback
        traceback.print_exc()
        return None


if __name__ == '__main__':
    file_path = 'TKB-20251-K66-69-du-kien-15.07.2025.xlsx'
    
    print("=" * 60)
    print("🔍 LỌC MÃ HỌC PHẦN ET VÀ EE")
    print("=" * 60)
    
    result = loc_ma_hoc_phan(file_path)
    
    if result is not None:
        print("\n" + "=" * 60)
        print("✅ Hoàn thành!")
        print("=" * 60)

