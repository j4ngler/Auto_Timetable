import pandas as pd
from pathlib import Path
import subprocess
import shutil
import argparse
import os

# Files
AI_RANK = Path('ai_ranked_classes.csv')
CLASSES = Path('classes_to_schedule.csv')
SLOTS = Path('timeslots.csv')
CONSTRAINTS = Path('constraints.json')
OUT = Path('schedule_final.csv')
TIMETABLE_ALL = Path('timetable_all.csv')


MAJOR = None  # EE/ET hoặc None

def ensure_ai_rank():
    if not AI_RANK.exists():
        print('Chua co ai_ranked_classes.csv — dang chay ai_recommender.py...')
        cmd = ['python', 'ai_recommender.py']
        if MAJOR in ('EE','ET'):
            cmd += ['--major', MAJOR]
        subprocess.run(cmd, check=False)


def reorder_classes_by_ai():
    if not AI_RANK.exists() or not CLASSES.exists():
        return
    ai = pd.read_csv(AI_RANK)
    cl = pd.read_csv(CLASSES) if CLASSES.exists() else pd.DataFrame()

    # Lọc theo ngành nếu có
    if MAJOR in ('EE','ET'):
        if 'CourseID' in ai.columns:
            ai = ai[ai['CourseID'].astype(str).str.startswith(MAJOR)]
        if not cl.empty and 'CourseID' in cl.columns:
            cl = cl[cl['CourseID'].astype(str).str.startswith(MAJOR)]

    # Loại bỏ các học phần không muốn xếp (ví dụ ETHICS)
    for df in (ai, cl):
        if 'CourseID' in df.columns:
            df.drop(df[df['CourseID'].astype(str).str.upper().str.startswith('ETHICS')].index, inplace=True)

    # Nếu sau khi lọc, CLASSES rỗng → tái tạo từ timetable_all.csv (theo major)
    if (cl is None or cl.empty) and TIMETABLE_ALL.exists():
        all_df = pd.read_csv(TIMETABLE_ALL)
        # Chuẩn hóa mã học phần
        if 'CourseID' not in all_df.columns and 'Mã_HP' in all_df.columns:
            all_df['CourseID'] = all_df['Mã_HP']
        if MAJOR in ('EE','ET') and 'CourseID' in all_df.columns:
            all_df = all_df[all_df['CourseID'].astype(str).str.startswith(MAJOR)]
        # Loại bỏ ETHICS
        if 'CourseID' in all_df.columns:
            all_df = all_df[~all_df['CourseID'].astype(str).str.upper().str.startswith('ETHICS')]
        # Lấy các trường cần cho classes_to_schedule
        subj_col = 'Tên_HP' if 'Tên_HP' in all_df.columns else 'SubjectName'
        room_col = 'Phòng' if 'Phòng' in all_df.columns else 'Room'
        dur_col = 'Buổi_số' if 'Buổi_số' in all_df.columns else 'Duration'
        tmp = all_df[['CourseID']].copy()
        tmp['SubjectName'] = all_df[subj_col] if subj_col in all_df.columns else ''
        tmp['Teacher'] = ''
        tmp['Duration'] = all_df[dur_col] if dur_col in all_df.columns else 1
        tmp['Capacity'] = ''
        tmp['RoomCandidates'] = all_df[room_col] if room_col in all_df.columns else ''
        tmp['Day'] = ''
        tmp['TimeSlot'] = ''
        tmp['RoomAssigned'] = ''
        # Tạo ClassID duy nhất theo thứ tự
        tmp['ClassID'] = [f"{cid}-{i+1}" for i, cid in enumerate(tmp['CourseID'])]
        cl = tmp[['ClassID','CourseID','SubjectName','Teacher','Duration','Capacity','RoomCandidates','Day','TimeSlot','RoomAssigned']]

    # Map điểm AI theo CourseID (ưu tiên cao xếp trước)
    ai_best = ai.groupby('CourseID', as_index=False)['ai_score'].max() if 'ai_score' in ai.columns else ai[['CourseID']].copy(); ai_best['ai_score']=0
    # Làm sạch mọi cột điểm AI thừa (ai_score, ai_score_x, ai_score_y, ...)
    cols_to_drop = [c for c in cl.columns if str(c).startswith('ai_score')]
    if cols_to_drop:
        cl = cl.drop(columns=cols_to_drop)
    merged = cl.merge(ai_best, on='CourseID', how='left', suffixes=('', '_ai'))
    if 'ai_score' not in merged.columns:
        merged['ai_score'] = 0
    merged['ai_score'] = merged['ai_score'].fillna(0)
    merged = merged.sort_values('ai_score', ascending=False)
    # Đảm bảo thứ tự cột ổn định, giữ nguyên các cột solver cần
    preferred_order = [
        'ClassID','CourseID','SubjectName','Teacher','Duration','Capacity','RoomCandidates',
        'Day','TimeSlot','RoomAssigned','ai_score'
    ]
    col_order = [c for c in preferred_order if c in merged.columns] + [c for c in merged.columns if c not in preferred_order]
    merged = merged[col_order]
    merged.to_csv(CLASSES, index=False, encoding='utf-8-sig')
    print('Da sap xep classes_to_schedule.csv theo AI score (major=%s)' % (MAJOR or 'ALL'))


def run_greedy():
    subprocess.run(['python', 'greedy_solver.py'], check=False)


def main():
    global MAJOR
    parser = argparse.ArgumentParser()
    parser.add_argument('--major', choices=['EE','ET'], default=None)
    args = parser.parse_args()
    MAJOR = args.major

    ensure_ai_rank()
    reorder_classes_by_ai()
    run_greedy()
    if OUT.exists():
        # Sao lưu theo major để dùng đa người dùng
        if MAJOR in ('EE','ET'):
            major_out = Path(f'schedule_final_{MAJOR}.csv')
            try:
                shutil.copyfile(OUT, major_out)
            except Exception:
                pass
        print(f'Pipeline hoan tat -> {OUT.resolve()}')


if __name__ == '__main__':
    main()
