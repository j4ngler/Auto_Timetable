import pandas as pd
from pathlib import Path

CLASSES = Path('classes_to_schedule.csv')
OUTPUT_ALL = Path('timetable_all.csv')


def main():
    if not CLASSES.exists():
        print(f'❌ Không tìm thấy {CLASSES.resolve()}')
        return

    dfc = pd.read_csv(CLASSES)
    if dfc.empty:
        print('⚠️ classes_to_schedule.csv rỗng')
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
    print(f'✅ Đã tạo {OUTPUT_ALL.resolve()} ({len(out)} dòng) từ classes_to_schedule.csv (header tiếng Việt)')


if __name__ == '__main__':
    main()
