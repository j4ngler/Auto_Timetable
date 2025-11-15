import pandas as pd
from pathlib import Path
import shutil
import os
import sys
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
import numpy as np

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
def get_path(filename):
    new_path = DATA_OUTPUT / filename
    if new_path.exists():
        return new_path
    old_path = PROJECT_ROOT / filename
    if old_path.exists():
        return old_path
    # Trả về đường dẫn mới dù chưa tồn tại (để tạo file mới)
    return new_path

TIMETABLE_ALL = get_path('timetable_all.csv')
TIMETABLE_USER = get_path('timetable_user.csv')
OUTPUT_RANK = get_path('ai_ranked_classes.csv')
CLASSES_CSV = get_path('classes_to_schedule.csv')

# Nếu chưa có timetable_all.csv, có thể sinh từ Ma_hoc_phan_ET_EE.xlsx bằng build_training_dataset.py


def parse_ranges(ranges_str):
    res = []
    if not isinstance(ranges_str, str) or not ranges_str.strip():
        return res
    for part in ranges_str.split(','):
        part = part.strip()
        if '-' in part:
            a, b = part.split('-', 1)
            res.append((a.strip(), b.strip()))
    return res


def time_in_ranges(timeslot, ranges):
    # timeslot dạng 'HH:MM-HH:MM' hoặc 'T1-3'
    if not isinstance(timeslot, str) or not timeslot:
        return 0
    if timeslot.startswith('T'):
        return 0  # bỏ qua dạng tiết nếu không có mapping
    try:
        start, end = timeslot.split('-', 1)
    except ValueError:
        return 0
    for a, b in ranges:
        if a <= start <= b:
            return 1
    return 0


def _split_clean(comma_str):
    if not isinstance(comma_str, str):
        return []
    return [t.strip() for t in comma_str.split(',') if t and t.strip()]


def build_training(df_all: pd.DataFrame, user_pref: pd.Series) -> pd.DataFrame:
    # Sinh nhãn mục tiêu (score) dựa trên sở thích người dùng — weak supervision
    pref_days = set(_split_clean(user_pref.get('PreferredDays', '')))
    pref_ranges = parse_ranges(user_pref.get('PreferredTimeSlots', ''))
    avoid_teachers = set(_split_clean(user_pref.get('AvoidTeachers', '')))
    preferred_teachers = set(_split_clean(user_pref.get('PreferredTeachers', '')))
    # Ưu tiên theo phòng/toà do người dùng chỉ định; nếu không có thì fallback theo bộ mặc định
    preferred_rooms = set(_split_clean(user_pref.get('PreferredRooms', '')))
    # Ưu tiên theo toà nhà (nếu có PreferredBuildings)
    preferred_buildings = set(_split_clean(user_pref.get('PreferredBuildings', '')))
    # Suy ra toà nhà từ danh sách phòng người dùng (tiền tố trước dấu '-')
    user_buildings_from_rooms = { r.split('-')[0].strip() for r in preferred_rooms if '-' in r }
    # Kết hợp preferred_buildings và user_buildings_from_rooms
    if preferred_buildings:
        user_buildings = preferred_buildings | user_buildings_from_rooms
    else:
        user_buildings = user_buildings_from_rooms
    # Fallback khi PreferredRooms rỗng
    default_buildings = { 'D3', 'C7', 'D3-5', 'D5', 'D7' }

    scores = []
    for _, r in df_all.iterrows():
        score = 0.0
        day = str(r.get('Day', ''))
        if day and day in pref_days:
            score += 1.0
        score += 1.0 * time_in_ranges(r.get('TimeSlot', ''), pref_ranges)
        # Chấm điểm theo phòng/toà
        room = str(r.get('Room', '') or r.get('RoomAssigned', '')).strip()
        building = room.split('-')[0].strip() if room else ''
        if room:
            # Nếu người dùng có PreferredRooms: ưu tiên khớp phòng, nếu không khớp phòng thì khớp theo toà
            if preferred_rooms:
                if room in preferred_rooms:
                    score += 1.0  # Khớp chính xác phòng
                elif building and building in user_buildings:
                    score += 0.5  # Khớp toà nhà
            elif preferred_buildings:
                # Chỉ có PreferredBuildings, không có PreferredRooms
                if building and building in preferred_buildings:
                    score += 0.8  # Ưu tiên toà nhà được chỉ định
                else:
                    score -= 0.3
            else:
                # Không cấu hình -> dùng bộ toà mặc định
                if building in default_buildings:
                    score += 1.0
                else:
                    score -= 0.5
        # Xử lý giáo viên: ưu tiên PreferredTeachers, tránh AvoidTeachers
        teacher = str(r.get('Teacher', '')).strip()
        if teacher:
            if teacher in preferred_teachers:
                score += 1.0  # Ưu tiên giáo viên được yêu thích
            elif teacher in avoid_teachers:
                score -= 1.0  # Tránh giáo viên không muốn
        scores.append(score)

    df_train = df_all.copy()
    df_train['score'] = scores
    return df_train


def train_and_rank(df_train: pd.DataFrame, df_all: pd.DataFrame) -> pd.DataFrame:
    # Bỏ Teacher; dùng Room như đặc trưng chính thay thế
    features = ['Day', 'TimeSlot', 'Room']
    target = 'score'

    df_train = df_train.fillna('')
    df_all = df_all.fillna('')

    pre = ColumnTransformer([
        ('cat', OneHotEncoder(handle_unknown='ignore'), features)
    ])

    model = Pipeline([
        ('pre', pre),
        ('rf', RandomForestRegressor(n_estimators=100, random_state=42))
    ])

    model.fit(df_train[features], df_train[target])
    preds = model.predict(df_all[features])

    out = df_all.copy()
    out['ai_score'] = preds
    out = out.sort_values('ai_score', ascending=False).reset_index(drop=True)
    return out


def main():
    # Cho phép lọc theo ngành qua tham số --major (EE/ET)
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--major', choices=['EE','ET'], default=None)
    args = parser.parse_args()

    # Debug: in đường dẫn để kiểm tra
    print(f'[INFO] Tim kiem timetable_all.csv tai: {TIMETABLE_ALL}')
    print(f'[INFO] Tim kiem timetable_user.csv tai: {TIMETABLE_USER}')
    print(f'[INFO] PROJECT_ROOT: {PROJECT_ROOT}')
    print(f'[INFO] DATA_OUTPUT: {DATA_OUTPUT}')
    
    if not TIMETABLE_ALL.exists() or not TIMETABLE_USER.exists():
        print(f'[ERROR] Thieu timetable_all.csv hoac timetable_user.csv.')
        print(f'[ERROR] timetable_all.csv exists: {TIMETABLE_ALL.exists()}')
        print(f'[ERROR] timetable_user.csv exists: {TIMETABLE_USER.exists()}')
        print(f'[ERROR] Hay chay build_training_dataset.py truoc.')
        return

    df_all = pd.read_csv(TIMETABLE_ALL)
    # Chuẩn hóa tên cột từ header tiếng Việt → cột đặc trưng nội bộ
    if 'Thứ' in df_all.columns:
        df_all['Day'] = df_all['Thứ']
    if 'Thời_gian' in df_all.columns:
        df_all['TimeSlot'] = df_all['Thời_gian']
    if 'Phòng' in df_all.columns:
        df_all['Room'] = df_all['Phòng']
    # Thêm ánh xạ mã học phần để merge với classes_to_schedule.csv
    if 'CourseID' not in df_all.columns and 'Mã_HP' in df_all.columns:
        df_all['CourseID'] = df_all['Mã_HP']

    # Lọc theo ngành nếu có
    if args.major in ('EE','ET'):
        if 'CourseID' in df_all.columns:
            df_all = df_all[df_all['CourseID'].astype(str).str.startswith(args.major)]
        elif 'Mã_HP' in df_all.columns:
            df_all = df_all[df_all['Mã_HP'].astype(str).str.startswith(args.major)]

    # Loại bỏ các học phần không mong muốn (ETHICS)
    if 'CourseID' in df_all.columns:
        df_all = df_all[~df_all['CourseID'].astype(str).str.upper().str.startswith('ETHICS')]
    elif 'Mã_HP' in df_all.columns:
        df_all = df_all[~df_all['Mã_HP'].astype(str).str.upper().str.startswith('ETHICS')]

    df_user = pd.read_csv(TIMETABLE_USER)
    user_pref = df_user.iloc[0] if len(df_user) > 0 else pd.Series()

    # Enrich dữ liệu nếu thiếu Day/TimeSlot/Teacher/Room bằng classes_to_schedule.csv
    if CLASSES_CSV.exists():
        cl = pd.read_csv(CLASSES_CSV)
        # Chỉ giữ các cột có thể bổ sung
        keep_cols = ['CourseID', 'Teacher', 'RoomCandidates', 'Day', 'TimeSlot']
        cl_small = cl[[c for c in keep_cols if c in cl.columns]].copy()
        # Bỏ merge nếu df_all chưa có CourseID
        if 'CourseID' in df_all.columns and 'CourseID' in cl_small.columns:
            df_all = df_all.merge(cl_small, on='CourseID', how='left', suffixes=('', '_cls'))
            for col in ['Teacher', 'Day', 'TimeSlot']:
                if f'{col}_cls' in df_all.columns:
                    df_all[col] = df_all[col].fillna(df_all[f'{col}_cls'])
            # Suy ra Room từ RoomCandidates nếu Room trống
            if 'Room' in df_all.columns and 'RoomCandidates' in df_all.columns:
                df_all['Room'] = df_all['Room'].fillna(df_all['RoomCandidates'].astype(str).str.split(',').str[0])
            # Xóa cột _cls tạm
            drop_cols = [c for c in df_all.columns if c.endswith('_cls')]
            df_all = df_all.drop(columns=drop_cols + ['RoomCandidates'], errors='ignore')

    df_train = build_training(df_all, user_pref)
    ranked = train_and_rank(df_train, df_all)

    ranked.to_csv(OUTPUT_RANK, index=False, encoding='utf-8-sig')
    # Lưu thêm file theo major để dùng lại khi đổi tài khoản
    if args.major in ('EE','ET'):
        try:
            major_output = DATA_OUTPUT / f'ai_ranked_classes_{args.major}.csv'
            shutil.copyfile(OUTPUT_RANK, major_output)
        except Exception:
            pass
    # Tránh ký tự unicode đặc biệt gây lỗi trên một số terminal Windows
    print(f'Da tao {OUTPUT_RANK.resolve()} (Top 10 hien thi mau)')
    
    # Tránh lỗi encode khi in tiếng Việt trên một số terminal Windows.
    # Bỏ in preview DataFrame; chỉ thông báo ngắn gọn ASCII.
    print('Preview top 10 duoc bo qua de tranh loi encoding tren Windows console.')


if __name__ == '__main__':
    main()
