import pandas as pd
import re
from pathlib import Path

# Ưu tiên dùng file đã được chèn header chuẩn nếu có
INPUT_FILE = Path('Ma_hoc_phan_ET_EE_fixed.xlsx') if Path('Ma_hoc_phan_ET_EE_fixed.xlsx').exists() else Path('Ma_hoc_phan_ET_EE.xlsx')
OUTPUT_ALL = Path('timetable_all.csv')
OUTPUT_USER = Path('timetable_user.csv')

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
    if not INPUT_FILE.exists():
        print(f'❌ Không thấy file {INPUT_FILE.resolve()}')
        return

    print('📖 Đọc dữ liệu...')

    # Nếu là file fixed (có header chuẩn), đọc trực tiếp với header=True
    if 'fixed' in str(INPUT_FILE):
        print('   📋 Phát hiện file có header chuẩn — đọc trực tiếp...')
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
        print('⚠️ File rỗng hoặc không đọc được')
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
    print(f'✅ Đã tạo {OUTPUT_ALL.resolve()} ({len(out)} dòng)')

    # Tạo file timetable_user.csv (bỏ AvoidTeachers theo yêu cầu)
    sample = pd.DataFrame({
        'PreferredDays': ['Mon,Tue,Thu'],
        'PreferredTimeSlots': ['07:00-11:00,13:00-17:00'],
        'PreferredRooms': [''],
        'MaxCredits': [20]
    })
    sample.to_csv(OUTPUT_USER, index=False, encoding='utf-8-sig')
    print(f'✅ Đã tạo {OUTPUT_USER.resolve()} (mẫu cấu hình ưu tiên người dùng)')


if __name__ == '__main__':
    main()
