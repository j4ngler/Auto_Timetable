import pandas as pd
import re
import json
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

# Đường dẫn file - ưu tiên vị trí mới, fallback vị trí cũ
def get_input_path(filename):
    new_path = DATA_INPUT / filename
    if new_path.exists():
        return new_path
    return PROJECT_ROOT / filename

def get_output_path(filename):
    return DATA_OUTPUT / filename

def get_config_path(filename):
    config_dir = PROJECT_ROOT / 'config'
    config_path = config_dir / filename
    if config_path.exists():
        return config_path
    # Fallback về thư mục gốc
    return PROJECT_ROOT / filename

# Ưu tiên dùng file đã được chèn header chuẩn nếu có
fixed_file = get_input_path('Ma_hoc_phan_ET_EE_fixed.xlsx')
normal_file = get_input_path('Ma_hoc_phan_ET_EE.xlsx')
INPUT_FILE = fixed_file if fixed_file.exists() else normal_file
OUTPUT_ALL = get_output_path('timetable_all.csv')
OUTPUT_USER = get_output_path('timetable_user.csv')
CONSTRAINTS_JSON = get_config_path('constraints.json')

# Các mẫu nhận diện cột
CODE_PATTERNS = [
    r'^m[aã] *h[oọ]c *ph[aă]n$',
    r'^m[aã] *_* *h[oọ]c *_* *ph[aă]n$',
    r'^(mã|ma)[ _]*hp$',
    r'^ma *_*hp$',
    r'^mã *_*hp$',
    r'^m[aã]_hp$',
    r'^mã hp$',
    r'^code$', r'^subject *code$', r'^(et|ee)[a-z0-9-]+$'
]
NAME_PATTERNS = [r'^(t[eê]n|name).*m[oô]n|subject *name|course *name']
TEACHER_PATTERNS = [r'^(gv|gi[aá]o *vi[eê]n|teacher)']
ROOM_PATTERNS = [r'^(ph[oò]ng|room)']
DAY_PATTERNS = [r'^(th[uứ]|ng[aà]y|day)']
TIME_PATTERNS = [r'^(ti[eê]t|ca|gi[oờ]|time|slot)']
CAP_PATTERNS = [r'^(s[iĩ]nh *vi[eê]n|s[cô] *t[iê]n|capacity|s[i]t)$']
DURATION_PATTERNS = [r'^(s[oố] *ti[eê]t|duration|ti[mê]n)$']
FACULTY_PATTERNS = [r'^(khoa|faculty)$']


def find_col(columns, regex_list):
    cols_norm = [str(c).strip().lower() for c in columns]
    for i, col in enumerate(cols_norm):
        for pat in regex_list:
            if re.search(pat, col):
                return columns[i]
    return None


def normalize_timeslot(value: str) -> str:
    if pd.isna(value):
        return ''
    s = str(value).strip()
    # Chuẩn hóa dạng HH:MM-HH:MM nếu có
    m = re.findall(r'(\d{1,2}[:h]\d{0,2}).{0,3}(\d{1,2}[:h]\d{0,2})', s)
    if m:
        a, b = m[0]
        a = a.replace('h', ':')
        b = b.replace('h', ':')
        if len(a) == 2:
            a = a + ':00'
        if len(b) == 2:
            b = b + ':00'
        if len(a) == 4:
            a = '0' + a
        if len(b) == 4:
            b = '0' + b
        return f'{a}-{b}'
    # Nếu là tiết: T1-3
    m2 = re.search(r'[tT](\d+)[ -–]+(\d+)', s)
    if m2:
        return f'T{m2.group(1)}-{m2.group(2)}'
    return s


def load_all_sheets(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    frames = []
    for sh in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sh)
        if df.empty:
            continue
        # loại bỏ các dòng header dài ở dòng 1 nếu chỉ có 1 cột nội dung
        if len(df.columns) == 1 and df.columns[0] == df.iloc[0, 0]:
            df = df.iloc[1:].reset_index(drop=True)
        df['__sheet__'] = sh
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


# Đưa hàm chuẩn hóa Day lên trước để dùng ở mọi nhánh

def normalize_day(v):
    s = str(v).strip().lower()
    mapping = {
        'thứ 2': 'Mon', 'thu 2': 'Mon', 't2': 'Mon', 'monday': 'Mon', 'mon': 'Mon', '2': 'Mon',
        'thứ 3': 'Tue', '3': 'Tue', 't3': 'Tue', 'tuesday': 'Tue', 'tue': 'Tue',
        'thứ 4': 'Wed', '4': 'Wed', 't4': 'Wed', 'wednesday': 'Wed', 'wed': 'Wed',
        'thứ 5': 'Thu', '5': 'Thu', 't5': 'Thu', 'thursday': 'Thu', 'thu': 'Thu',
        'thứ 6': 'Fri', '6': 'Fri', 't6': 'Fri', 'friday': 'Fri', 'fri': 'Fri',
        'thứ 7': 'Sat', '7': 'Sat', 't7': 'Sat', 'saturday': 'Sat', 'sat': 'Sat',
        'chủ nhật': 'Sun', 'cn': 'Sun', 'sunday': 'Sun', 'sun': 'Sun'
    }
    return mapping.get(s, v)


def normalize_day_general(v):
    s = str(v).strip().lower()
    if s.isdigit():
        mapping_num = {
            '2': 'Mon', '3': 'Tue', '4': 'Wed', '5': 'Thu', '6': 'Fri', '7': 'Sat', '8': 'Sun'
        }
        return mapping_num.get(s, s)
    return normalize_day(v)


def main():
    import argparse
    parser = argparse.ArgumentParser(description='Xử lý dữ liệu Excel và tạo dataset')
    parser.add_argument('--force', action='store_true', help='Ghi đè file timetable_user.csv nếu đã tồn tại')
    args = parser.parse_args()
    
    if not INPUT_FILE.exists():
        print(f'[ERROR] Khong thay file {INPUT_FILE.resolve()}')
        return

    print('[INFO] Doc du lieu...')

    # Nếu là file fixed (có header chuẩn), đọc trực tiếp với header=True
    if 'fixed' in str(INPUT_FILE):
        print('   [INFO] Phat hien file co header chuan — doc truc tiep...')
        xls = pd.ExcelFile(INPUT_FILE)
        all_dfs = []
        for sh in xls.sheet_names:
            df_sh = pd.read_excel(INPUT_FILE, sheet_name=sh)
            if not df_sh.empty:
                all_dfs.append(df_sh)
        raw = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
    else:
        raw = load_all_sheets(INPUT_FILE)

    if raw.empty:
        print('[WARNING] File rong hoac khong doc duoc')
        return

    # Chỉ giữ các cột yêu cầu
    desired_cols = ['Kỳ','Trường_Viện_Khoa','Mã_lớp','Mã_lớp_kèm','Mã_HP','Tên_HP','Khối_lượng','Ghi_chú','Buổi_số','Thứ','Thời_gian','BĐ','KT','Kíp','Tuần','Phòng']
    for c in desired_cols:
        if c not in raw.columns:
            raw[c] = ''
    out = raw[desired_cols].copy()

    # Fallback: điền 'Mã_HP' nếu trống bằng cách quét toàn bộ ô trong dòng để tìm mã ET/EE
    code_regex = re.compile(r'(?i)\b((?:ET|EE)[A-Z0-9-]+)\b')
    def extract_code_row(row):
        mhp = str(row.get('Mã_HP', '')).strip()
        m0 = code_regex.search(mhp)
        if m0:
            return m0.group(1).upper()
        for v in row.values:
            m = code_regex.search(str(v) if v is not None else '')
            if m:
                return m.group(1).upper()
        return ''

    out['Mã_HP'] = out.apply(extract_code_row, axis=1)

    # Chuẩn hóa Mã_HP: bỏ ký tự lạ, upper, strip
    out['Mã_HP'] = out['Mã_HP'].astype(str).str.replace(r'[^A-Za-z0-9-]', '', regex=True).str.upper().str.strip()

    # Lọc: giữ các dòng có Mã_HP bắt đầu bằng ET/EE (không dùng regex group để tránh warning)
    mask = out['Mã_HP'].str.upper().str.startswith(('ET', 'EE'))
    out = out[mask]

    # Chuẩn hóa Thứ và Thời_gian
    out['Thứ'] = out['Thứ'].apply(normalize_day_general)
    out['Thời_gian'] = out['Thời_gian'].apply(normalize_timeslot)

    # Xuất
    out = out.replace({pd.NA: '', None: ''}).fillna('')
    out.to_csv(OUTPUT_ALL, index=False, encoding='utf-8-sig')
    print(f'[SUCCESS] Da tao {OUTPUT_ALL.resolve()} ({len(out)} dong)')

    # Tạo file timetable_user.csv với đầy đủ đặc trưng
    # Đọc từ constraints.json nếu có, nếu không dùng giá trị mặc định
    default_prefs = {
        'PreferredDays': 'Mon,Tue,Wed,Thu,Fri,Sat',
        'PreferredTimeSlots': '07:00-09:00,09:00-11:00,13:00-15:00,15:00-17:00',
        'PreferredRooms': 'D3-504,D3-505,C7-205,C7-206,D5-101,D5-102',
        'MaxCredits': 24,
        'MinCredits': 18,
        'PreferredTeachers': '',
        'AvoidTeachers': '',
        'PreferredBuildings': 'D3,C7,D5,D7'
    }
    
    # Đọc từ constraints.json
    if CONSTRAINTS_JSON.exists():
        try:
            with open(CONSTRAINTS_JSON, 'r', encoding='utf-8') as f:
                constraints = json.load(f)
            
            # Đọc credits
            credits_config = constraints.get('credits', {})
            if credits_config.get('min_total') is not None:
                default_prefs['MinCredits'] = int(credits_config['min_total'])
            if credits_config.get('max_total') is not None:
                default_prefs['MaxCredits'] = int(credits_config['max_total'])
            
            # Đọc buildings
            buildings_config = constraints.get('buildings', {})
            preferred_buildings = buildings_config.get('preferred', [])
            if preferred_buildings:
                default_prefs['PreferredBuildings'] = ','.join(preferred_buildings)
            
            # Đọc rooms
            rooms_config = constraints.get('rooms', {})
            preferred_rooms = rooms_config.get('preferred', [])
            if preferred_rooms:
                default_prefs['PreferredRooms'] = ','.join(preferred_rooms)
            
            # Đọc time_slots
            time_slots_config = constraints.get('time_slots', {})
            preferred_slots = time_slots_config.get('preferred', [])
            if preferred_slots:
                default_prefs['PreferredTimeSlots'] = ','.join(preferred_slots)
            elif time_slots_config.get('preferred_morning'):
                default_prefs['PreferredTimeSlots'] = '07:00-09:00,09:00-11:00'
            elif time_slots_config.get('preferred_afternoon'):
                default_prefs['PreferredTimeSlots'] = '13:00-15:00,15:00-17:00'
            
            # Đọc priority Day để suy ra PreferredDays
            priority_config = constraints.get('priority', {})
            preferred_days = priority_config.get('Day', [])
            if preferred_days:
                default_prefs['PreferredDays'] = ','.join(preferred_days)
            
            # Đọc teachers
            teachers_config = constraints.get('teachers', {})
            preferred_teachers = teachers_config.get('preferred', [])
            avoid_teachers = teachers_config.get('avoid', [])
            if preferred_teachers:
                default_prefs['PreferredTeachers'] = ','.join(preferred_teachers)
            if avoid_teachers:
                default_prefs['AvoidTeachers'] = ','.join(avoid_teachers)
                
        except Exception as e:
            print(f'[WARNING] Khong doc duoc constraints.json: {e}. Su dung gia tri mac dinh.')
    
    # Tạo DataFrame từ preferences
    sample = pd.DataFrame({
        'PreferredDays': [default_prefs['PreferredDays']],
        'PreferredTimeSlots': [default_prefs['PreferredTimeSlots']],
        'PreferredRooms': [default_prefs['PreferredRooms']],
        'MaxCredits': [default_prefs['MaxCredits']],
        'MinCredits': [default_prefs['MinCredits']],
        'PreferredTeachers': [default_prefs['PreferredTeachers']],
        'AvoidTeachers': [default_prefs['AvoidTeachers']],
        'PreferredBuildings': [default_prefs['PreferredBuildings']]
    })
    
    # Ghi file timetable_user.csv
    if not OUTPUT_USER.exists() or args.force:
        sample.to_csv(OUTPUT_USER, index=False, encoding='utf-8-sig')
        if args.force:
            print(f'[SUCCESS] Da cap nhat {OUTPUT_USER.resolve()} (dong bo voi constraints.json)')
        else:
            print(f'[SUCCESS] Da tao {OUTPUT_USER.resolve()} (dong bo voi constraints.json)')
    else:
        print(f'[INFO] File {OUTPUT_USER.resolve()} da ton tai. Bo qua tao moi.')
        print(f'[INFO] De cap nhat, chay voi flag --force: python build_training_dataset.py --force')


if __name__ == '__main__':
    main()
