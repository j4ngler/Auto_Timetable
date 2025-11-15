import pandas as pd
import json
from pathlib import Path
import re

# Đổi sang đọc trực tiếp file Excel đã lọc
INPUT_XLSX = Path('Ma_hoc_phan_ET_EE_fixed.xlsx')
OUT_CLASSES = Path('classes_to_schedule.csv')
OUT_SLOTS = Path('timeslots.csv')
OUT_CONSTRAINTS = Path('constraints.json')

# Khung timeslot mặc định (có thể chỉnh)
DEFAULT_SLOTS = [
    {'Slot': 1, 'Start': '07:00', 'End': '09:00'},
    {'Slot': 2, 'Start': '09:00', 'End': '11:00'},
    {'Slot': 3, 'Start': '13:00', 'End': '15:00'},
    {'Slot': 4, 'Start': '15:00', 'End': '17:00'},
]
DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']


def normalize_room_candidates(room_val: str) -> str:
    if not isinstance(room_val, str) or not room_val.strip():
        return ''
    # Tách theo , ; / khoảng trắng
    parts = re.split(r'[;,\/\s]+', room_val)
    parts = [p for p in parts if p]
    return ','.join(sorted(set(parts)))


def find_col(columns, regex_list):
    cols_norm = [str(c).strip().lower() for c in columns]
    for i, col in enumerate(cols_norm):
        for pat in regex_list:
            if re.search(pat, col):
                return columns[i]
    return None


CODE_PATTERNS = [r'^m[aã] *h[oọ]c *ph[aă]n$', r'^m[aã] *hp$', r'^code$', r'^subject *code$', r'^(et|ee)[a-z0-9-]+$']
NAME_PATTERNS = [r'(t[eê]n|name).*m[oô]n|subject *name|course *name']
TEACHER_PATTERNS = [r'^(gv|gi[aá]o *vi[eê]n|teacher)']
ROOM_PATTERNS = [r'^(ph[oò]ng|room)']
DURATION_PATTERNS = [r'^(bu[oố]i|ti[eê]t|duration)$']


def load_excel_any(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    frames = []
    for sh in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sh)
        if df.empty:
            continue
        df['__sheet__'] = sh
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def main():
    if not INPUT_XLSX.exists():
        print(f'❌ Không thấy {INPUT_XLSX.resolve()} — hãy đặt file xlsx đã lọc cùng thư mục')
        return

    raw = load_excel_any(INPUT_XLSX)
    if raw.empty:
        print('⚠️ File Excel rỗng hoặc không đọc được')
        return

    # Dò cột chính - ƯU TIÊN header tiếng Việt chuẩn
    cols = list(raw.columns)
    vn = {str(c).strip(): c for c in cols}

    # Cột mã học phần
    col_code = vn.get('Mã_HP') or vn.get('Mã HP') or find_col(cols, CODE_PATTERNS) or cols[0]
    # Cột tên học phần
    col_name = vn.get('Tên_HP') or vn.get('Tên HP') or find_col(cols, NAME_PATTERNS)
    # Cột giảng viên
    col_teacher = vn.get('Giảng_viên') or vn.get('Giảng viên') or find_col(cols, TEACHER_PATTERNS)
    # Cột phòng
    col_room = vn.get('Phòng') or find_col(cols, ROOM_PATTERNS)
    # Cột số buổi (duration)
    col_duration = vn.get('Buổi_số') or vn.get('Buổi số') or find_col(cols, DURATION_PATTERNS)

    # Trích code ET/EE từ dữ liệu thô (ưu tiên cột mã). Nếu không có, quét toàn bộ dòng
    code_regex = re.compile(r'\b((?:ET|EE)[A-Z0-9-]+)\b', re.IGNORECASE)
    course_ids = []
    for _, row in raw.iterrows():
        code = None
        # 1) từ cột mã
        if col_code in row:
            m = code_regex.search(str(row[col_code]))
            if m:
                code = m.group(1).upper()
        # 2) fallback: quét mọi ô trong dòng
        if code is None:
            for val in row.values:
                m = code_regex.search(str(val))
                if m:
                    code = m.group(1).upper()
                    break
        course_ids.append(code)

    df = pd.DataFrame({'CourseID': course_ids})
    df = df[df['CourseID'].notna()]

    # Tạo classes_to_schedule.csv — biến cần tìm: Day, TimeSlot, RoomAssigned
    classes = pd.DataFrame()
    classes['ClassID'] = [f"{cid}-{i+1}" for i, cid in enumerate(df['CourseID'])]
    classes['CourseID'] = df['CourseID']
    classes['SubjectName'] = raw[col_name] if col_name in raw.columns else ''
    classes['Teacher'] = raw[col_teacher] if col_teacher in raw.columns else ''
    classes['Duration'] = raw[col_duration] if col_duration in raw.columns else 3
    classes['Capacity'] = ''
    classes['RoomCandidates'] = (raw[col_room].apply(normalize_room_candidates) if col_room in raw.columns else '')

    # Các cột để solver điền
    classes['Day'] = ''
    classes['TimeSlot'] = ''
    classes['RoomAssigned'] = ''

    classes.to_csv(OUT_CLASSES, index=False, encoding='utf-8-sig')
    print(f'✅ Đã tạo {OUT_CLASSES.resolve()} ({len(classes)} dòng)')

    # Tạo timeslots.csv (cartesian DAYS x DEFAULT_SLOTS)
    ts_rows = []
    for d in DAYS:
        for s in DEFAULT_SLOTS:
            ts_rows.append({'Day': d, **s})
    slots = pd.DataFrame(ts_rows)
    slots.to_csv(OUT_SLOTS, index=False, encoding='utf-8-sig')
    print(f'✅ Đã tạo {OUT_SLOTS.resolve()} ({len(slots)} slots)')

    # Tạo constraints.json (ràng buộc cơ bản)
    constraints = {
        'no_overlap': {
            'by': ['Teacher', 'RoomAssigned'],
            'message': 'Không trùng giáo viên/phòng trong cùng Day+TimeSlot'
        },
        'room_candidates': True,
        'max_classes_per_slot': None,
        'priority': {
            'Day': ['Mon', 'Tue', 'Thu', 'Wed', 'Fri', 'Sat'],
            'TimeSlot': [1, 2, 3, 4]
        }
    }
    OUT_CONSTRAINTS.write_text(json.dumps(constraints, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'✅ Đã tạo {OUT_CONSTRAINTS.resolve()}')

    print('\n📌 Gợi ý tiếp theo:')
    print('- Dùng OR-Tools/Pulp để viết solver đọc các file trên và xuất lịch tối ưu')
    print('- Hoặc viết greedy baseline: xếp lần lượt từng lớp theo ưu tiên, tránh xung đột')


if __name__ == '__main__':
    main()
