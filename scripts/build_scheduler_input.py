import pandas as pd
import json
from pathlib import Path
import re

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
CONFIG_DIR = PROJECT_ROOT / 'config'

# Đường dẫn file - ưu tiên vị trí mới, fallback vị trí cũ
def get_input_path(filename):
    new_path = DATA_INPUT / filename
    if new_path.exists():
        return new_path
    return PROJECT_ROOT / filename

def get_output_path(filename):
    return DATA_OUTPUT / filename

def get_config_path(filename):
    return CONFIG_DIR / filename

# Đổi sang đọc trực tiếp file Excel đã lọc
INPUT_XLSX = get_input_path('Ma_hoc_phan_ET_EE_fixed.xlsx')
OUT_CLASSES = get_output_path('classes_to_schedule.csv')
OUT_SLOTS = get_output_path('timeslots.csv')
OUT_CONSTRAINTS = get_config_path('constraints.json')

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
        print(f'[ERROR] Khong thay {INPUT_XLSX.resolve()} — hay dat file xlsx da loc cung thu muc')
        return

    raw = load_excel_any(INPUT_XLSX)
    if raw.empty:
        print('[WARNING] File Excel rong hoac khong doc duoc')
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
    print(f'[SUCCESS] Da tao {OUT_CLASSES.resolve()} ({len(classes)} dong)')

    # Tạo timeslots.csv (cartesian DAYS x DEFAULT_SLOTS)
    ts_rows = []
    for d in DAYS:
        for s in DEFAULT_SLOTS:
            ts_rows.append({'Day': d, **s})
    slots = pd.DataFrame(ts_rows)
    slots.to_csv(OUT_SLOTS, index=False, encoding='utf-8-sig')
    print(f'[SUCCESS] Da tao {OUT_SLOTS.resolve()} ({len(slots)} slots)')

    # Tạo constraints.json (ràng buộc cơ bản) - cấu trúc mới v2.0
    constraints = {
        'no_overlap': {
            'by': ['Teacher', 'RoomAssigned'],
            'message': 'Không trùng giáo viên/phòng trong cùng Day+TimeSlot',
            'strict': True
        },
        'room_candidates': True,
        'max_classes_per_slot': None,
        'max_classes_per_day': None,
        'min_gap_between_classes': 0,
        'priority': {
            'Day': ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
            'TimeSlot': [1, 2, 3, 4],
            'description': 'Thứ tự ưu tiên khi sắp xếp: Day trước, sau đó TimeSlot'
        },
        'user_preferences': {
            'enabled': True,
            'source': 'timetable_user.csv',
            'description': 'Tích hợp sở thích người dùng từ timetable_user.csv'
        },
        'credits': {
            'max_per_day': None,
            'max_total': None,
            'min_total': None,
            'description': 'Ràng buộc về tín chỉ (nếu null thì không giới hạn)'
        },
        'teachers': {
            'preferred': [],
            'avoid': [],
            'description': 'Danh sách giáo viên ưa thích/tránh (đọc từ timetable_user.csv)'
        },
        'buildings': {
            'preferred': [],
            'description': 'Danh sách toà nhà ưa thích (đọc từ timetable_user.csv)'
        },
        'rooms': {
            'preferred': [],
            'description': 'Danh sách phòng ưa thích (đọc từ timetable_user.csv)'
        },
        'time_slots': {
            'preferred': [],
            'avoid_rush_hours': False,
            'preferred_morning': False,
            'preferred_afternoon': False,
            'description': 'Khung giờ ưa thích (đọc từ timetable_user.csv)'
        },
        'fallback': {
            'allow_teacher_overlap': False,
            'allow_room_overlap': False,
            'description': 'Khi không thể gán slot hợp lệ, có cho phép vi phạm ràng buộc không'
        },
        'version': '2.0',
        'last_updated': '2025-01-11',
        'notes': 'Cập nhật để hỗ trợ các đặc trưng mới: PreferredTeachers, AvoidTeachers, PreferredBuildings, MaxCredits, MinCredits'
    }
    OUT_CONSTRAINTS.write_text(json.dumps(constraints, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'[SUCCESS] Da tao {OUT_CONSTRAINTS.resolve()}')

    print('\n[INFO] Goi y tiep theo:')
    print('- Dung OR-Tools/Pulp de viet solver doc cac file tren va xuat lich toi uu')
    print('- Hoac viet greedy baseline: xep lan luot tung lop theo uu tien, tranh xung dot')


if __name__ == '__main__':
    main()
