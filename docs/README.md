# SPCN Timetable (Flask)

Ứng dụng web quản lý – xếp thời khoá biểu (TKB) theo ngành (EE/ET) với gợi ý AI, chạy cục bộ bằng Flask.

## 1) Yêu cầu môi trường
- Windows 10/11, Python 3.12 (khuyến nghị dùng venv riêng `venv312`).
- Những thư mục/fiIe dữ liệu chính nằm ở thư mục gốc dự án:
  - `timetable_all.csv`: dữ liệu lớp toàn trường đã chuẩn hoá.
  - `timetable_user.csv`: cấu hình sở thích người dùng.
  - `constraints.json`: ràng buộc xếp lịch.

## 2) Cài đặt nhanh
```bash
# 1) Tạo venv và kích hoạt
python -m venv web\venv
web\venv\Scripts\activate

# 2) Cài thư viện
pip install -r requirements.txt  # nếu có

# 3) Chạy web
python web\app.py
# Trình duyệt sẽ mở tại http://127.0.0.1:5000/
```

Nếu bạn có nhiều phiên bản Python, bảo đảm chạy đúng bản 3.12 (ví dụ thông qua `py -3.12`).

## 3) Đăng ký/Đăng nhập
- Truy cập trang chủ → nếu chưa đăng nhập sẽ thấy trang chào mừng với 2 nút: Đăng nhập / Tạo tài khoản.
- Tạo tài khoản: nhập email, username, mật khẩu, chọn chế độ (student/teacher) và ngành `EE` hoặc `ET`.
- Sau khi đăng nhập, tên người dùng (góc phải trên) dẫn đến trang Hồ sơ để cập nhật thông tin (tên, điện thoại, khoá, lớp, avatar, ngành).

## 4) Điều hướng chính trong Sidebar
- `Dashboard`: số liệu tổng quan; lịch hôm nay (cá nhân); top gợi ý AI; thao tác nhanh.
- `TKB cá nhân`: xem TKB dạng lưới (ngày × ca) từ `schedule_recommended.csv` (hoặc `TKB_ca_nhan.csv`).
- `TKB toàn trường`: xem lưới TKB toàn trường từ `schedule_final*.csv`.
- `Tạo TKB`: giao diện chạy xếp lịch và xem kết quả.
- `Upload`: tải lên/ghi đè các file CSV/JSON hệ thống.
- `Constraints`: chỉnh ràng buộc bằng form trực quan (không cần nhập JSON).
- `Hồ sơ`: cập nhật thông tin, chuyển ngành EE/ET.

## 5) Luồng công việc phổ biến
### 5.1 Chạy gợi ý AI và xếp lịch
- Từ Dashboard hoặc trang `Tạo TKB`:
  - Bấm “Chạy gợi ý lớp” để sinh `ai_ranked_classes_{major}.csv`.
  - Bấm “Chạy sắp xếp TKB” để tạo `schedule_final_{major}.csv`.
  - Bấm “Tạo TKB gợi ý” để sinh `schedule_recommended.csv` cho cá nhân.
- Hệ thống tự nhận ngành từ phiên đăng nhập. Mọi kết quả được cache theo ngành: `..._EE.csv` và `..._ET.csv`.

### 5.2 Tùy biến sở thích cá nhân (ảnh hưởng AI)
- Mở `timetable_user.csv` (hoặc Upload file này):
  - `PreferredDays`: danh sách ngày, ví dụ `Mon,Wed,Fri`.
  - `PreferredTimeSlots`: `HH:MM-HH:MM` cách nhau bởi dấu phẩy, ví dụ `07:00-09:00,13:00-15:00`.
  - `PreferredRooms`: danh sách phòng/toà, ví dụ `D3-504,C7-205,D5`.
  - `MaxCredits`: tín chỉ mong muốn tối đa (dùng ở bước gợi ý cá nhân).

### 5.3 Chỉnh ràng buộc xếp lịch
- Vào `Constraints` → bật/tắt các tuỳ chọn:
  - Không trùng trong cùng ca: `Teacher`, `RoomAssigned`.
  - `room_candidates`: chỉ dùng phòng trong RoomCandidates.
  - `max_classes_per_slot`: giới hạn số lớp/slot (để trống = không giới hạn).
  - Thứ tự ưu tiên `Day` và `TimeSlot`.
- Bấm “Lưu” để ghi `constraints.json`, sau đó chạy lại “Chạy sắp xếp TKB”.

## 6) AI hoạt động như thế nào (tóm tắt)
- Mô hình: RandomForestRegressor (rừng cây hồi quy) học ánh xạ từ `(Day, TimeSlot, Room)` → `ai_score`.
- Nhãn học (weak supervision) được tính từ sở thích người dùng:
  - +1 nếu `Day` thuộc `PreferredDays`.
  - +1 nếu `TimeSlot` thuộc `PreferredTimeSlots`.
  - +1 nếu `Room` nằm trong `PreferredRooms`; +0.5 nếu cùng toà; nếu không cấu hình phòng thì ưu tiên các toà `{D3, C7, D3-5, D5, D7}`, còn lại −0.5.
- Dự đoán `ai_score` cho toàn bộ lớp (lọc theo ngành EE/ET) → xếp hạng → dẫn dắt bộ xếp lịch.

## 7) Cơ chế theo ngành (EE/ET)
- Khi đăng nhập hoặc đổi ngành, hệ thống “làm nóng” cache nền (nếu thiếu) để tạo `ai_ranked_classes_{major}.csv` và `schedule_final_{major}.csv`.
- Dashboard/Preview/Download luôn ưu tiên đọc theo ngành của phiên đăng nhập. Nếu chưa có file theo ngành, hệ thống tạm dùng file chung và sẽ tạo bản theo ngành ở nền.

## 8) File kết quả/ý nghĩa
- `ai_ranked_classes.csv` / `ai_ranked_classes_{EE|ET}.csv`: xếp hạng lớp theo điểm AI.
- `classes_to_schedule.csv`: danh sách lớp đầu vào (đã sắp theo AI) cho solver.
- `schedule_final.csv` / `schedule_final_{EE|ET}.csv`: TKB toàn trường (đã gán Day/Slot/Room).
- `schedule_recommended.csv`: lịch cá nhân sau tối ưu với các phạt mềm.

## 9) Sự cố thường gặp
- Lỗi Unicode khi chạy script trên Windows console: các script đã chuyển sang ASCII; nếu còn lỗi, chạy từ giao diện web để xem log.
- Không thấy số liệu sau khi đổi ngành: đợi vài giây để cache nền hoàn tất, hoặc bấm “Chạy sắp xếp TKB”.
- Thiếu phòng/slot khiến nhiều dòng trống: vào `Constraints` và/hoặc cập nhật `RoomCandidates`/`timeslots.csv`.

## 10) Cấu trúc thư mục chính
```
web/
  app.py                 # Flask app, route, job runner, SQLite user
  templates/             # Giao diện Jinja2 (dashboard, auth, TKB,...)
  static/                # CSS, ảnh
ai_recommender.py        # Huấn luyện + suy luận AI → ai_ranked_classes*.csv
run_pipeline.py          # Orchestrate: AI rank → reorder → greedy_solver
greedy_solver.py         # Gán Day/Slot/Room theo ràng buộc
recommend_schedule.py    # Chọn TKB cá nhân với phạt mềm
```

## 11) Bảo mật & phạm vi
- Ứng dụng mặc định chạy `127.0.0.1:5000` (cục bộ). Không khuyến nghị mở rộng ra mạng ngoài nếu chưa cấu hình xác thực/HTTPS.

---
Bất kỳ góp ý hoặc lỗi phát sinh, vui lòng tạo issue/note kèm log ở trang `Tạo TKB` → “Xem kết quả/Poll” để mình hỗ trợ nhanh.

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
