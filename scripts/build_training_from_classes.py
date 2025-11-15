import pandas as pd
from pathlib import Path

# Tìm thư mục gốc dự án
def get_project_root():
    """Tìm thư mục gốc dự án"""
    current = Path(__file__).resolve()
    if current.parent.name == 'scripts':
        return current.parent.parent  # Lên 2 cấp: scripts -> project_root
    return Path.cwd()

PROJECT_ROOT = get_project_root()
DATA_OUTPUT = PROJECT_ROOT / 'data' / 'output'

# Đường dẫn file - ưu tiên vị trí mới, fallback vị trí cũ
def get_data_path(filename):
    new_path = DATA_OUTPUT / filename
    if new_path.exists():
        return new_path
    return PROJECT_ROOT / filename

CLASSES = get_data_path('classes_to_schedule.csv')
OUTPUT_ALL = get_data_path('timetable_all.csv')


def main():
    if not CLASSES.exists():
        print(f'[ERROR] Khong tim thay {CLASSES.resolve()}')
        return

    dfc = pd.read_csv(CLASSES)
    if dfc.empty:
        print('[WARNING] classes_to_schedule.csv rong')
        return

    # Chuẩn hóa sang header tiếng Việt chuẩn của TKB
    def first_room(val):
        if isinstance(val, str) and val:
            return val.split(',')[0].strip()
        return ''

    slot_map = {1: '07:00-09:00', 2: '09:00-11:00', 3: '13:00-15:00', 4: '15:00-17:00'}
    def map_slot(v):
        try:
            i = int(v)
            return slot_map.get(i, '')
        except Exception:
            return str(v) if isinstance(v, str) else ''

    # Xử lý Phòng: ưu tiên RoomAssigned, nếu không có thì lấy phòng đầu tiên từ RoomCandidates
    if 'RoomAssigned' in dfc.columns:
        room_col = dfc['RoomAssigned'].fillna('')
    else:
        room_col = pd.Series([''] * len(dfc))
    
    if 'RoomCandidates' in dfc.columns:
        candidates_room = dfc['RoomCandidates'].apply(first_room)
        # Điền phòng từ candidates chỉ khi RoomAssigned trống
        mask_empty = (room_col == '') | room_col.isna()
        room_col = room_col.where(~mask_empty, candidates_room).fillna('')

    out = pd.DataFrame({
        'Kỳ': '',
        'Trường_Viện_Khoa': 'Electrical & Electronics',
        'Mã_lớp': dfc.get('ClassID', ''),
        'Mã_lớp_kèm': '',
        'Mã_HP': dfc.get('CourseID', ''),
        'Tên_HP': dfc.get('SubjectName', ''),
        'Tên_HP_Tiếng_Anh': '',
        'Khối_lượng': '',
        'Ghi_chú': '',
        'Buổi_số': dfc.get('Duration', ''),
        'Thứ': dfc.get('Day', ''),
        'Thời_gian': dfc.get('TimeSlot', '').apply(map_slot),
        'BĐ': '',
        'KT': '',
        'Kíp': '',
        'Tuần': '',
        'Phòng': room_col,
        'Cần_TN': '',
        'SLĐK': '',
        'SL_Max': dfc.get('Capacity', ''),
        'Trạng_thái': '',
        'Loại_lớp': '',
        'Đợt_mở': '',
        'Mã_QL': '',
        'Hệ': '',
        'TeachingType': '',
        'mainclass': '',
        'Sessionid': '',
        'Statusid': '',
        'Khóa': ''
    })

    out.to_csv(OUTPUT_ALL, index=False, encoding='utf-8-sig')
    print(f'[SUCCESS] Da tao {OUTPUT_ALL.resolve()} ({len(out)} dong) tu classes_to_schedule.csv (header tieng Viet)')


if __name__ == '__main__':
    main()
