import pandas as pd
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
DATA_OUTPUT = PROJECT_ROOT / 'data' / 'output'

# Đường dẫn file - ưu tiên vị trí mới, fallback vị trí cũ
def get_data_path(filename):
    new_path = DATA_OUTPUT / filename
    if new_path.exists():
        return new_path
    return PROJECT_ROOT / filename

AI_RANK = get_data_path('ai_ranked_classes.csv')   # đầu vào đã có cột ai_score
OUT_REC = get_data_path('schedule_recommended.csv')

# đặt mục tiêu tín chỉ cho kỳ tới
MIN_CREDITS = 14
MAX_CREDITS = 20
# Số môn tối đa (yêu cầu 9–10)
MAX_COURSES = 10
# Phạt mềm khi xếp nhiều môn trong 1 ngày hoặc các môn sát nhau
MIN_GAP_MINUTES = 30
DAILY_MAX_SOFT = 5

def parse_credits(kl):
    # Khối_lượng dạng: 3(3-0-1-6) hoặc '3' → lấy số đầu
    if pd.isna(kl): 
        return 0
    m = re.match(r'\s*(\d+)', str(kl))
    return int(m.group(1)) if m else 0

def times_overlap(a, b):
    # a,b: 'HH:MM-HH:MM' hoặc '0920-1145'
    pat = r"(\d{2})(\d{2})-(\d{2})(\d{2})"
    def get_int_time(s):
        m = re.match(pat, s.replace(':',''))
        if m:
            h1,m1,h2,m2 = map(int, m.groups())
            return h1*60+m1, h2*60+m2
        # backup cho HH:MM-HH:MM
        try:
            sa,ea = s.split('-')
            ha,ma = map(int, sa.split(':'))
            hb,mb = map(int, ea.split(':'))
            return ha*60+ma, hb*60+mb
        except:
            return 0,0
    s1,e1 = get_int_time(a)
    s2,e2 = get_int_time(b)
    return not (e1 <= s2 or e2 <= s1)

def slot_gap_minutes(a, b):
    pat = r"(\d{2})(\d{2})-(\d{2})(\d{2})"
    def parse(s):
        m = re.match(pat, s.replace(':',''))
        if m:
            h1,m1,h2,m2 = map(int, m.groups())
            return h1*60+m1, h2*60+m2
        try:
            sa,ea = s.split('-')
            ha,ma = map(int, sa.split(':'))
            hb,mb = map(int, ea.split(':'))
            return ha*60+ma, hb*60+mb
        except:
            return 0,0
    s1,e1 = parse(a)
    s2,e2 = parse(b)
    if e1 <= s2:
        return s2 - e1
    if e2 <= s1:
        return s1 - e2
    return 0

def main():
    df = pd.read_csv(AI_RANK).fillna('')
    # Chuẩn hoá tên cột
    if 'Day' not in df and 'Thứ' in df: df['Day'] = df['Thứ']
    if 'TimeSlot' not in df and 'Thời_gian' in df: df['TimeSlot'] = df['Thời_gian']
    if 'CourseID' not in df and 'Mã_HP' in df: df['CourseID'] = df['Mã_HP']
    if 'SubjectName' not in df and 'Tên_HP' in df: df['SubjectName'] = df['Tên_HP']
    if 'Khối_lượng' not in df.columns: df['Khối_lượng'] = ''
    if 'Loại_lớp' not in df.columns: df['Loại_lớp'] = ''
    if 'ai_score' not in df.columns:
        raise SystemExit('Thiếu cột ai_score trong ai_ranked_classes.csv (hãy chạy ai_recommender.py trước).')
    df = df.sort_values('ai_score', ascending=False).reset_index(drop=True)
    df['credits'] = df['Khối_lượng'].apply(parse_credits)

    chosen = []
    taken_theory = set()
    taken_lab = set()
    total_credits = 0
    total_courses = 0

    for _, r in df.iterrows():
        day = str(r['Day']).strip()
        ts = str(r['TimeSlot']).strip()
        cid = str(r['CourseID']).strip()
        typ = str(r['Loại_lớp']).upper().replace(' ','')
        cr = int(r['credits'])
        if total_credits >= MAX_CREDITS:
            break
        is_lab = ('TN' in typ) or ('THUC HANH' in typ) or ('THỰC HÀNH' in typ)
        if is_lab and cid in taken_lab:
            continue
        if (not is_lab) and cid in taken_theory:
            continue
        conflict = False
        for x in chosen:
            if day == x['Day'] and times_overlap(ts, x['TimeSlot']):
                conflict = True
                break
        if conflict:
            continue
        # Phạt mềm theo mật độ trong ngày và khoảng cách giờ
        same_day = [x for x in chosen if x['Day'] == day]
        penalty = 0.0
        if len(same_day) >= DAILY_MAX_SOFT:
            penalty += 0.5 + 0.2 * (len(same_day) - DAILY_MAX_SOFT)
        if same_day:
            min_gap = min(slot_gap_minutes(ts, x['TimeSlot']) for x in same_day)
            if min_gap < MIN_GAP_MINUTES:
                penalty += 0.5
        eff_score = float(r['ai_score']) - penalty
        if eff_score < 0.5:
            continue
        chosen.append({
            'CourseID': cid,
            'SubjectName': r.get('SubjectName',''),
            'Day': day,
            'TimeSlot': ts,
            'Room': r.get('Room',''),
            'Loại_lớp': r.get('Loại_lớp',''),
            'Khối_lượng': r.get('Khối_lượng',''),
            'credits': cr,
            'ai_score': r['ai_score']
        })
        total_credits += cr
        total_courses += 1
        if is_lab:
            taken_lab.add(cid)
        else:
            taken_theory.add(cid)
        if total_credits >= MIN_CREDITS and total_courses >= MAX_COURSES:
            break

    out = pd.DataFrame(chosen)
    out.to_csv(OUT_REC, index=False, encoding='utf-8-sig')
    # In ASCII để tránh lỗi encoding trên Windows console
    print(f'Done. Chon {len(out)} lop, tong {out.credits.sum()} TC -> {OUT_REC}')

if __name__ == '__main__':
    main()
