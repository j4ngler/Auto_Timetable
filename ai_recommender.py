import pandas as pd
from pathlib import Path
import shutil
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
import numpy as np

TIMETABLE_ALL = Path('timetable_all.csv')
TIMETABLE_USER = Path('timetable_user.csv')
OUTPUT_RANK = Path('ai_ranked_classes.csv')
CLASSES_CSV = Path('classes_to_schedule.csv')

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
    # Ưu tiên theo phòng/toà do người dùng chỉ định; nếu không có thì fallback theo bộ mặc định
    preferred_rooms = set(_split_clean(user_pref.get('PreferredRooms', '')))
    # Suy ra toà nhà từ danh sách phòng người dùng (tiền tố trước dấu '-')
    user_buildings = { r.split('-')[0].strip() for r in preferred_rooms if '-' in r }
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
                    score += 1.0
                elif building and building in user_buildings:
                    score += 0.5
            else:
                # Không cấu hình -> dùng bộ toà mặc định
                if building in default_buildings:
                    score += 1.0
                else:
                    score -= 0.5
        # Nếu có cấu hình tránh giáo viên thì vẫn áp dụng khi dữ liệu có Teacher
        teacher = str(r.get('Teacher', '')).strip()
        if teacher and teacher in avoid_teachers:
            score -= 1.0
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

    if not TIMETABLE_ALL.exists() or not TIMETABLE_USER.exists():
        print('❌ Thiếu timetable_all.csv hoặc timetable_user.csv. Hãy chạy build_training_dataset.py trước.')
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
            Path(f'ai_ranked_classes_{args.major}.csv').write_text('')
            shutil.copyfile(OUTPUT_RANK, Path(f'ai_ranked_classes_{args.major}.csv'))
        except Exception:
            pass
    # Tránh ký tự unicode đặc biệt gây lỗi trên một số terminal Windows
    print(f'Da tao {OUTPUT_RANK.resolve()} (Top 10 hien thi mau)')
    
    # Tránh lỗi encode khi in tiếng Việt trên một số terminal Windows.
    # Bỏ in preview DataFrame; chỉ thông báo ngắn gọn ASCII.
    print('Preview top 10 duoc bo qua de tranh loi encoding tren Windows console.')


if __name__ == '__main__':
    main()
