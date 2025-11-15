"""
Script lọc mã học phần ET và EE từ file Excel TKB
"""

import pandas as pd
import sys
import re
from pathlib import Path

# Tìm thư mục gốc dự án
def get_project_root():
    """Tìm thư mục gốc dự án"""
    current = Path(__file__).resolve()
    if current.parent.name == 'scripts':
        return current.parent.parent  # Lên 2 cấp: scripts -> project_root
    return Path.cwd()

PROJECT_ROOT = get_project_root()
DATA_INPUT = PROJECT_ROOT / 'data' / 'input'
DATA_OUTPUT = PROJECT_ROOT / 'data' / 'output'


def loc_ma_hoc_phan(file_path):
    """
    Lọc các mã học phần bắt đầu bằng ET hoặc EE
    """
    try:
        # Đọc tất cả sheets từ Excel
        print("[INFO] Dang doc file Excel...")
        excel_file = pd.ExcelFile(file_path)
        
        print(f"[SUCCESS] Tim thay {len(excel_file.sheet_names)} sheet(s)")
        
        # Lưu tất cả kết quả (dòng chứa ET/EE) và tất cả mã tìm được
        all_rows = []
        all_codes = set()
        
        # Duyệt qua từng sheet
        for sheet_name in excel_file.sheet_names:
            print(f"\n[INFO] Dang xu ly sheet: '{sheet_name}'...")
            
            # Đọc sheet (không dùng header để tránh dòng tiêu đề dài)
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str)
            
            # In thử vài dòng đầu để xem cấu trúc
            if len(df) > 0:
                print(f"   [INFO] Sheet co {len(df)} dong, {len(df.columns)} cot")
                # Regex bắt mã HP: bắt đầu bằng ET/EE, theo sau là chữ số/chữ in/ dấu gạch
                code_pattern = re.compile(r"\b(ET|EE)[A-Z0-9-]+\b", re.IGNORECASE)

                # Tạo mask dòng nào có chứa mã ET/EE ở bất kỳ cột nào
                row_has_code = df.apply(
                    lambda row: any(bool(code_pattern.search(str(val))) for val in row.values), axis=1
                )

                matched_rows = df[row_has_code].copy()

                if len(matched_rows) > 0:
                    print(f"   [SUCCESS] Tim thay {len(matched_rows)} dong chua ma ET/EE")

                    # Trích xuất mã từ toàn bộ sheet để tổng hợp danh sách mã duy nhất
                    for val in df.astype(str).values.flatten():
                        for m in code_pattern.findall(str(val)):
                            all_codes.add(m.upper())

                    # Gắn tên sheet để truy vết
                    matched_rows['Sheet'] = sheet_name
                    all_rows.append(matched_rows)
                else:
                    print("   [WARNING] Khong tim thay ma ET/EE trong sheet nay")
        
        # Gộp tất cả dòng khớp
        if all_rows:
            result_df = pd.concat(all_rows, ignore_index=True)
            
            # Lưu ra file Excel (bản gốc, giữ nguyên cấu trúc ô)
            output_file = DATA_INPUT / 'Ma_hoc_phan_ET_EE.xlsx'
            result_df.to_excel(output_file, index=False, header=True)
            print(f"\n[SUCCESS] Da loc duoc {len(result_df)} dong")
            print(f"[INFO] Da luu vao file: {output_file}")

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

            fixed_file = DATA_INPUT / 'Ma_hoc_phan_ET_EE_fixed.xlsx'
            # Ghi đảm bảo có hàng tiêu đề
            fixed_df.to_excel(fixed_file, index=False, header=True)
            print(f"[INFO] Dong thoi tao: {fixed_file} (da chen hang tieu de cot)")

            # In danh sách các mã học phần duy nhất (từ regex)
            unique_codes = sorted(all_codes)
            print(f"\n[INFO] Danh sach ma hoc phan ET/EE ({len(unique_codes)} ma):")
            for code in unique_codes:
                print(f"   - {code}")
            
            # Lưu danh sách mã vào file text
            txt_file = DATA_OUTPUT / 'Danh_sach_ma_ET_EE.txt'
            with open(txt_file, 'w', encoding='utf-8') as f:
                f.write("Danh sách mã học phần ET và EE\n")
                f.write("=" * 50 + "\n\n")
                for code in unique_codes:
                    f.write(f"{code}\n")
            print(f"[INFO] Da luu danh sach ma vao: {txt_file}")
            
            return fixed_df
        else:
            print("\n[WARNING] Khong tim thay ma hoc phan ET hoac EE nao!")
            return None
            
    except Exception as e:
        print(f"[ERROR] Loi: {e}")
        import traceback
        traceback.print_exc()
        return None


if __name__ == '__main__':
    # Tìm file input ở data/input/ hoặc thư mục gốc
    input_file = DATA_INPUT / 'TKB-20251-K66-69-du-kien-15.07.2025.xlsx'
    if not input_file.exists():
        input_file = PROJECT_ROOT / 'TKB-20251-K66-69-du-kien-15.07.2025.xlsx'
    
    print("=" * 60)
    print("LOC MA HOC PHAN ET VA EE")
    print("=" * 60)
    
    result = loc_ma_hoc_phan(str(input_file))
    
    if result is not None:
        print("\n" + "=" * 60)
        print("[SUCCESS] Hoan thanh!")
        print("=" * 60)

