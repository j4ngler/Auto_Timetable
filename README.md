# 📚 SPCN_HaiAnh - Xử lý TKB ET/EE

> **Xem hướng dẫn chi tiết pipeline**: `PIPELINE_HUONG_DAN.md` ⭐

## Tóm Tắt Nhanh

### Pipeline Hoàn Chỉnh (1 Lệnh):
```bash
python loc_ma_hoc_phan.py          # Lọc mã ET/EE
python build_training_dataset.py    # Tạo dataset
python build_scheduler_input.py     # Tạo input solver
python run_pipeline.py              # AI + Greedy → schedule_final.csv
```

### Kết Quả:
- `schedule_final.csv` - Thời khóa biểu đã xếp tự động

---

## 1) AI gợi ý lớp học (Training Dataset)

B1. Đặt file `Ma_hoc_phan_ET_EE.xlsx` vào thư mục này (đã lọc ET/EE).

B2. Tạo dataset huấn luyện:
```bash
python build_training_dataset.py
```
Sinh ra:
- `timetable_all.csv` — dữ liệu chuẩn hóa: CourseID, SubjectName, Teacher, Room, Day, TimeSlot, Duration, Capacity, Faculty
- `timetable_user.csv` — file cấu hình ưu tiên người dùng (mẫu)

Gợi ý huấn luyện Random Forest:
- Trích đặc trưng từ Day/TimeSlot/Teacher/Room → mã hóa one-hot
- Mục tiêu: dự đoán lớp phù hợp theo ưu tiên người dùng

## 2) Auto Scheduler (Constraint Solver)

B1. Chạy bước 1 để có `timetable_all.csv`.

B2. Tạo input cho solver:
```bash
python build_scheduler_input.py
```
Sinh ra:
- `classes_to_schedule.csv` — danh sách lớp cần xếp; solver sẽ gán Day/TimeSlot/RoomAssigned
- `timeslots.csv` — lưới ngày/khung giờ chuẩn
- `constraints.json` — ràng buộc cơ bản (không trùng giáo viên/phòng)

B3. Viết solver (khuyến nghị OR-Tools):
- Đọc `classes_to_schedule.csv`, `timeslots.csv`, `constraints.json`
- Biến quyết định: (class, day, slot, room)
- Ràng buộc: không trùng giáo viên/phòng cùng (day, slot), tôn trọng RoomCandidates

## Lưu ý
- Dữ liệu gốc có thể thiếu cột; script sẽ suy luận hoặc để trống hợp lý.
- Có thể sửa danh sách DAYS/DEFAULT_SLOTS trong `build_scheduler_input.py` cho phù hợp thực tế.
