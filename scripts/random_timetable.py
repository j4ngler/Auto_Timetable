import pandas as pd
import random
from datetime import timedelta, datetime
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

monhoc_df = pd.read_csv(get_data_path('timetable_all.csv'))
# Lọc các môn có mã học phần bắt đầu bằng 'ET'
et_mask = monhoc_df['Mã_HP'].astype(str).str.startswith('ET')
monhoc_et = monhoc_df[et_mask].drop_duplicates(subset=['Tên_HP'])

TIET_MIN = 45
BREAK_MIN = 10
PERIODS = 6
WEEKDAYS = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6']
SESSIONS = ['Sáng', 'Chiều']


def get_time_blocks(session):
    if session == 'Sáng':
        start = datetime.strptime('07:00', '%H:%M')
    else:
        start = datetime.strptime('13:00', '%H:%M')
    blocks = []
    curr = start
    for i in range(PERIODS):
        end = curr + timedelta(minutes=TIET_MIN)
        blocks.append((curr.strftime("%H:%M"), end.strftime("%H:%M")))
        if i < PERIODS-1:
            curr = end + timedelta(minutes=BREAK_MIN)
        else:
            curr = end
    return blocks

def generate_full_timetable(monhoc_et):
    timetable = []
    for weekday in WEEKDAYS:
        n_buoi = random.choice([1, 2])
        sessions_today = random.sample(SESSIONS, k=n_buoi)
        for buoi in sessions_today:
            num_periods = 4 if buoi == 'Sáng' else random.choice([3, 4])
            slots = sorted(random.sample(range(PERIODS), num_periods))
            num_subjects = random.choice([2, 3]) if num_periods > 2 else 2
            monhoc_list = monhoc_et['Tên_HP'].tolist()
            subjects = random.sample(monhoc_list, num_subjects)
            tiethoc_left = num_periods
            split_periods = []
            for i in range(num_subjects - 1):
                max_periods = min(3, tiethoc_left - (num_subjects - i - 1))
                min_periods = 1
                p = random.randint(min_periods, max_periods)
                split_periods.append(p)
                tiethoc_left -= p
            split_periods.append(tiethoc_left)
            cur = 0
            for i, subject in enumerate(subjects):
                info = monhoc_et[monhoc_et['Tên_HP'] == subject].iloc[0]
                for _ in range(split_periods[i]):
                    tiet_idx = slots[cur]
                    tgbd, tgkt = get_time_blocks(buoi)[tiet_idx]
                    # Thời gian học theo đúng context tiết (Tiết số, khung giờ)
                    timetable.append({
                        'Kỳ': info['Kỳ'],
                        'Trường_Viện_Khoa': info['Trường_Viện_Khoa'],
                        'Mã_lớp': info['Mã_lớp'],
                        'Mã_HP': info['Mã_HP'],
                        'Tên_HP': info['Tên_HP'],
                        'Khối_lượng': info['Khối_lượng'],
                        'Buổi_số': buoi,
                        'Thứ': weekday,
                        'Thời_gian': f"{tgbd}-{tgkt}",
                        'Kíp': tiet_idx+1,
                        'Phòng': info['Phòng'] if 'Phòng' in info else ''
                    })
                    cur += 1
    return sorted(timetable, key=lambda x: (WEEKDAYS.index(x['Thứ']), SESSIONS.index(x['Buổi_số']), x['Kíp']))

def main():
    tbl = generate_full_timetable(monhoc_et)
    df = pd.DataFrame(tbl)
    # sắp xếp đúng thứ tự theo trường tiêu đề gốc
    col_order = ['Kỳ','Trường_Viện_Khoa','Mã_lớp','Mã_HP','Tên_HP',
                 'Khối_lượng', 'Buổi_số','Thứ','Thời_gian','Kíp','Phòng']
    df = df[col_order]
    print(df.head())
    output_file = get_data_path('TKB_ca_nhan.csv')
    df.to_csv(output_file, index=False)
    print('Đã lưu file: TKB_ca_nhan.csv')

if __name__ == '__main__':
    main()