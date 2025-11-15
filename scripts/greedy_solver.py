import pandas as pd
import json
from pathlib import Path
import csv

# Tìm thư mục gốc dự án
def get_project_root():
    """Tìm thư mục gốc dự án"""
    current = Path(__file__).resolve()
    if current.parent.name == 'scripts':
        return current.parent.parent  # Lên 2 cấp: scripts -> project_root
    return Path.cwd()

PROJECT_ROOT = get_project_root()
DATA_OUTPUT = PROJECT_ROOT / 'data' / 'output'
CONFIG_DIR = PROJECT_ROOT / 'config'

# Đường dẫn file - ưu tiên vị trí mới, fallback vị trí cũ
def get_data_path(filename):
    new_path = DATA_OUTPUT / filename
    if new_path.exists():
        return new_path
    return PROJECT_ROOT / filename

def get_config_path(filename):
    new_path = CONFIG_DIR / filename
    if new_path.exists():
        return new_path
    return PROJECT_ROOT / filename

CLASSES_CSV = get_data_path('classes_to_schedule.csv')
SLOTS_CSV = get_data_path('timeslots.csv')
CONSTRAINTS_JSON = get_config_path('constraints.json')
OUTPUT_SCHEDULE = get_data_path('schedule_final.csv')


def load_inputs():
    classes = pd.read_csv(CLASSES_CSV)
    slots = pd.read_csv(SLOTS_CSV)
    constraints = json.loads(CONSTRAINTS_JSON.read_text(encoding='utf-8'))
    return classes, slots, constraints


def build_slot_priority(slots: pd.DataFrame, constraints: dict):
    # Sắp xếp theo Day, Slot theo ưu tiên trong constraints
    day_order = constraints.get('priority', {}).get('Day', [])
    slot_order = constraints.get('priority', {}).get('TimeSlot', [])

    def day_key(d):
        d = str(d)
        return day_order.index(d) if d in day_order else len(day_order)

    def slot_key(s):
        try:
            s = int(s)
        except Exception:
            s = 999
        return slot_order.index(s) if s in slot_order else len(slot_order)

    slots_sorted = slots.copy()
    slots_sorted['__day_ord__'] = slots_sorted['Day'].map(day_key)
    slots_sorted['__slot_ord__'] = slots_sorted['Slot'].map(slot_key)
    slots_sorted = slots_sorted.sort_values(['__day_ord__', '__slot_ord__']).reset_index(drop=True)
    return slots_sorted[['Day', 'Slot']]


def extract_room_candidates(val) -> list:
    if not isinstance(val, str) or not val.strip():
        return []
    return [x for x in str(val).split(',') if x]


def choose_room(available_rooms: set, candidates: list, used_rooms_in_slot: set):
    # Ưu tiên room trong candidates chưa bị dùng ở slot
    for r in candidates:
        if r not in used_rooms_in_slot:
            return r
    # Nếu không có, chọn một phòng chưa dùng từ tập available_rooms (phòng thật)
    for r in sorted(available_rooms):
        if r not in used_rooms_in_slot:
            return r
    # Nếu vẫn không có, trả về chuỗi rỗng (để caller thử slot khác)
    return ''


def greedy_schedule(classes: pd.DataFrame, slots: pd.DataFrame, constraints: dict) -> pd.DataFrame:
    # Chuẩn bị cấu trúc theo dõi xung đột: (Day, Slot) -> set(Teacher), set(Room)
    teacher_used = {}  # key: (Day, Slot) -> set
    room_used = {}     # key: (Day, Slot) -> set

    # Tập phòng tổng hợp từ candidates + phòng lấy từ timetable_all.csv (nếu có)
    all_rooms = set()
    classes['RoomCandidates'] = classes['RoomCandidates'].fillna('')
    for rc in classes['RoomCandidates']:
        all_rooms.update(extract_room_candidates(rc))

    # Bổ sung phòng từ timetable_all.csv
    tt_path = get_data_path('timetable_all.csv')
    if tt_path.exists():
        try:
            with tt_path.open('r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    room = (row.get('Phòng') or '').strip()
                    if room:
                        all_rooms.add(room)
        except Exception:
            pass

    # Sắp xếp lớp theo CourseID/Teacher để ổn định (có thể thay đổi chiến lược)
    classes_sorted = classes.copy()
    classes_sorted = classes_sorted.sort_values(['Teacher', 'CourseID']).reset_index(drop=True)

    # Duyệt từng lớp và gán slot đầu tiên hợp lệ theo ưu tiên
    assignments = []
    for _, row in classes_sorted.iterrows():
        teacher = str(row.get('Teacher', '')).strip()
        candidates = extract_room_candidates(row.get('RoomCandidates', ''))
        assigned = False

        for _, slot in slots.iterrows():
            day = slot['Day']
            ts = slot['Slot']
            key = (day, ts)

            # Constraint: không trùng giáo viên cùng (Day, Slot)
            if teacher:
                if key in teacher_used and teacher in teacher_used[key]:
                    continue

            # Chọn phòng hợp lệ
            used_rooms = room_used.get(key, set())
            room = choose_room(all_rooms, candidates, used_rooms)
            if not room or room in used_rooms:
                continue

            # Gán
            assignments.append({
                'ClassID': row['ClassID'],
                'CourseID': row['CourseID'],
                'SubjectName': row.get('SubjectName', ''),
                'Day': day,
                'TimeSlot': ts,
                'RoomAssigned': room,
                'Duration': row.get('Duration', 3),
                'Capacity': row.get('Capacity', '')
            })

            # Cập nhật used sets
            teacher_used.setdefault(key, set()).add(teacher)
            room_used.setdefault(key, set()).add(room)
            assigned = True
            break

        if not assigned:
            # Ép gán: bỏ qua ràng buộc giáo viên, tìm slot đầu tiên còn chỗ và cấp phòng dự phòng nếu cần
            for _, slot in slots.iterrows():
                day = slot['Day']
                ts = slot['Slot']
                key = (day, ts)
                used_rooms = room_used.get(key, set())
                room = choose_room(all_rooms, candidates, used_rooms)
                if not room or room in used_rooms:
                    continue

                assignments.append({
                    'ClassID': row['ClassID'],
                    'CourseID': row['CourseID'],
                    'SubjectName': row.get('SubjectName', ''),
                    'Day': day,
                    'TimeSlot': ts,
                    'RoomAssigned': room,
                    'Duration': row.get('Duration', 3),
                    'Capacity': row.get('Capacity', '')
                })
                room_used.setdefault(key, set()).add(room)
                assigned = True
                break

            # Trường hợp bất khả kháng (không nên xảy ra): vẫn ghi để theo dõi
            if not assigned:
                assignments.append({
                    'ClassID': row['ClassID'],
                    'CourseID': row['CourseID'],
                    'SubjectName': row.get('SubjectName', ''),
                    'Day': '',
                    'TimeSlot': '',
                    'RoomAssigned': '',
                    'Duration': row.get('Duration', 3),
                    'Capacity': row.get('Capacity', '')
                })

    return pd.DataFrame(assignments)


def main():
    if not CLASSES_CSV.exists() or not SLOTS_CSV.exists() or not CONSTRAINTS_JSON.exists():
        print('[ERROR] Thieu file dau vao. Can classes_to_schedule.csv, timeslots.csv, constraints.json')
        return

    classes, slots, constraints = load_inputs()
    slots_priority = build_slot_priority(slots, constraints)
    schedule = greedy_schedule(classes, slots_priority, constraints)
    schedule.to_csv(OUTPUT_SCHEDULE, index=False, encoding='utf-8-sig')
    print(f'[SUCCESS] Da tao {OUTPUT_SCHEDULE.resolve()} ({len(schedule)} dong)')


if __name__ == '__main__':
    main()
