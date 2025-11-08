# 🔄 Hướng Dẫn Pipeline - Chạy Các File Theo Thứ Tự

> Hướng dẫn chi tiết cách chạy từng bước trong dự án

---

## 📋 Tổng Quan Pipeline

### 🎯 2 Hướng Chính:

1. **AI Gợi Ý Lớp Học** → Mô hình Random Forest gợi ý lớp phù hợp
2. **Auto Scheduler** → Tự động xếp thời khóa biểu

### 🔗 Pipeline Kết Hợp

```
TKB Gốc → Lọc ET/EE → Dataset → AI Ranking → Greedy Solver → Lịch Cuối
```

---

## 🗺️ Sơ Đồ Pipeline Chi Tiết

```
┌─────────────────────────────────────────────────────────────────┐
│                    BƯỚC 1: CHUẨN BỊ DỮ LIỆU                     │
└─────────────────────────────────────────────────────────────────┘
                              ↓
        TKB-20251-K66-69-du-kien-15.07.2025.xlsx (File gốc)
                              ↓
                    python loc_ma_hoc_phan.py
                              ↓
        ┌─────────────────────┴─────────────────────┐
        ↓                                           ↓
Ma_hoc_phan_ET_EE.xlsx            Ma_hoc_phan_ET_EE_fixed.xlsx
(Bản gốc)                         (Có header chuẩn) ⭐
        │                         │
        │                         │
        └─────────────┬───────────┘
                      ↓
        python build_training_dataset.py
                      ↓
┌─────────────────────────────────────────────────┐
│  timetable_all.csv (626 dòng - header tiếng Việt) │
│  timetable_user.csv (cấu hình ưu tiên)           │
└─────────────────────────────────────────────────┘


┌─────────────────────────────────────────────────────────────────┐
│              BƯỚC 2A: AI GỢI Ý LỚP HỌC (Tùy Chọn)               │
└─────────────────────────────────────────────────────────────────┘
                              ↓
                    timetable_all.csv
                    timetable_user.csv
                              ↓
                    python ai_recommender.py
                              ↓
            ┌───────────────────────────────┐
            │ ai_ranked_classes.csv          │
            │ (Các lớp được xếp hạng theo AI) │
            └───────────────────────────────┘


┌─────────────────────────────────────────────────────────────────┐
│            BƯỚC 2B: AUTO SCHEDULER (Xếp TKB Tự Động)            │
└─────────────────────────────────────────────────────────────────┘
                              ↓
        Ma_hoc_phan_ET_EE_fixed.xlsx
        (Hoặc Ma_hoc_phan_ET_EE.xlsx)
                              ↓
            python build_scheduler_input.py
                              ↓
    ┌──────────────┬───────────────┬─────────────────┐
    ↓              ↓               ↓                 ↓
classes_to_schedule.csv   timeslots.csv   constraints.json
(644 dòng - chưa có      (24 slots)      (Ràng buộc)
 Day/TimeSlot/Room)                      
                              ↓
                    python greedy_solver.py
                              ↓
            ┌───────────────────────────────┐
            │ schedule_final.csv             │
            │ (TKB đã xếp tự động)           │
            └───────────────────────────────┘


┌─────────────────────────────────────────────────────────────────┐
│         BƯỚC 3: PIPELINE KẾT HỢP (AI + Greedy) ⭐               │
└─────────────────────────────────────────────────────────────────┘
                              ↓
                    python run_pipeline.py
                              ↓
    (Tự động gọi: ai_recommender.py → greedy_solver.py)
                              ↓
            ┌───────────────────────────────┐
            │ schedule_final.csv             │
            │ (Vừa hợp lệ, vừa hợp gu người dùng) │
            └───────────────────────────────┘
```

---

## 📝 Chi Tiết Từng Bước

### **BƯỚC 1: Lọc Mã Học Phần ET/EE**

**Mục đích**: Lọc ra các lớp học phần bắt đầu bằng ET hoặc EE

```bash
cd SPCN_HaiAnh
python loc_ma_hoc_phan.py
```

**Input**: `TKB-20251-K66-69-du-kien-15.07.2025.xlsx`

**Output**:
- ✅ `Ma_hoc_phan_ET_EE.xlsx` - File lọc (giữ nguyên cấu trúc)
- ✅ `Ma_hoc_phan_ET_EE_fixed.xlsx` - File có header chuẩn ⭐
- ✅ `Danh_sach_ma_ET_EE.txt` - Danh sách mã học phần

---

### **BƯỚC 2A: Tạo Dataset Huấn Luyện (Cho AI)**

**Mục đích**: Chuẩn hóa dữ liệu để huấn luyện mô hình AI

```bash
python build_training_dataset.py
```

**Input**: `Ma_hoc_phan_ET_EE_fixed.xlsx` (ưu tiên) hoặc `Ma_hoc_phan_ET_EE.xlsx`

**Output**:
- ✅ `timetable_all.csv` - Dataset đầy đủ (626 dòng, 30 cột header tiếng Việt)
- ✅ `timetable_user.csv` - Cấu hình ưu tiên người dùng

**Lưu ý**: 
- File này có đầy đủ thông tin từ Excel gốc
- Header: Kỳ, Trường_Viện_Khoa, Mã_lớp, Mã_HP, Tên_HP, Thứ, Thời_gian, Phòng...

---

### **BƯỚC 2B: Tạo Input Cho Scheduler**

**Mục đích**: Tạo file đầu vào cho solver tự động xếp lịch

```bash
python build_scheduler_input.py
```

**Input**: `Ma_hoc_phan_ET_EE_fixed.xlsx` hoặc `Ma_hoc_phan_ET_EE.xlsx`

**Output**:
- ✅ `classes_to_schedule.csv` - Danh sách lớp cần xếp (644 dòng)
  - Cột `Day`, `TimeSlot`, `RoomAssigned` để trống (solver sẽ điền)
- ✅ `timeslots.csv` - Lưới ngày/khung giờ (24 slots = 6 ngày × 4 ca)
- ✅ `constraints.json` - Ràng buộc giao dịch (không trùng giáo viên/phòng)

---

### **BƯỚC 3A: AI Gợi Ý (Tùy Chọn)**

**Mục đích**: Sử dụng AI để xếp hạng lớp theo ưu tiên người dùng

```bash
python ai_recommender.py
```

**Input**:
- `timetable_all.csv`
- `timetable_user.csv`

**Output**:
- ✅ `ai_ranked_classes.csv` - Các lớp được xếp hạng theo điểm AI

**Cách hoạt động**:
- Đọc sở thích từ `timetable_user.csv` (PreferredDays, PreferredTimeSlots...)
- Tính điểm cho từng lớp
- Sắp xếp theo điểm cao → thấp

#### Giải thích cách tính điểm (ai_score)
- File liên quan: `ai_recommender.py` (hàm `build_training` + `train_and_rank`)
- Đặc trưng đưa vào model: `Day`, `TimeSlot`, `Teacher`, `Room` (được one-hot)
- Nhãn huấn luyện (score) tạo theo “sở thích” người dùng (weak-supervision):
  - `+1` nếu `Day ∈ PreferredDays`
  - `+1` nếu `TimeSlot ∈ PreferredTimeSlots` (so sánh theo khoảng giờ HH:MM-HH:MM)
  - `-1` nếu `Teacher ∈ AvoidTeachers`
- Mô hình: `RandomForestRegressor` học từ (features → score) và dự đoán `ai_score` cho mọi lớp.
- Lưu ý:
  - Nếu `timetable_all.csv` thiếu `Day/TimeSlot/Teacher` → điểm khó phân hóa (thường ~0 hoặc hằng số).
  - Có thể điều chỉnh logic chấm điểm trong `ai_recommender.py` để tăng phân hóa (ví dụ: không khớp Day/Time → `-0.5`, khớp PreferredRooms → `+0.3`).

#### Cách cải thiện ai_score
- Thu hẹp `PreferredDays` / `PreferredTimeSlots` để tạo chênh lệch rõ ràng.
- Điền `AvoidTeachers` / `PreferredRooms` nếu có.
- Đảm bảo `timetable_all.csv` có `Thứ`, `Thời_gian`, `Phòng` (nên tạo từ Excel gốc).

---

### **BƯỚC 3B: Greedy Solver (Xếp TKB Tự Động)**

**Mục đích**: Tự động xếp thời khóa biểu hợp lệ

```bash
python greedy_solver.py
```

**Input**:
- `classes_to_schedule.csv`
- `timeslots.csv`
- `constraints.json`

**Output**:
- ✅ `schedule_final.csv` - TKB đã xếp tự động

**Cách hoạt động**:
- Duyệt từng lớp, gán slot đầu tiên hợp lệ
- Tránh xung đột: không trùng giáo viên/phòng cùng slot
- Ưu tiên slot theo thứ tự trong `constraints.json`

---

### **BƯỚC 3C: Pipeline Kết Hợp (Khuyến Nghị) ⭐**

**Mục đích**: Kết hợp AI ranking + Greedy solver

```bash
python run_pipeline.py
```

**Luồng hoạt động**:
1. Kiểm tra `ai_ranked_classes.csv` → Nếu chưa có, tự động gọi `ai_recommender.py`
2. Sắp xếp `classes_to_schedule.csv` theo điểm AI (cao → thấp)
3. Gọi `greedy_solver.py` để xếp lịch
4. Xuất `schedule_final.csv` - Vừa hợp lệ, vừa hợp gu!

---

## 🚀 Pipeline Hoàn Chỉnh (1 Lệnh)

Nếu muốn chạy từ đầu đến cuối:

```bash
cd SPCN_HaiAnh

# Bước 1: Lọc dữ liệu
python loc_ma_hoc_phan.py

# Bước 2: Tạo dataset + input cho solver
python build_training_dataset.py
python build_scheduler_input.py

# Bước 3: Chạy pipeline kết hợp
python run_pipeline.py
```

**Kết quả**: `schedule_final.csv` - Thời khóa biểu đã xếp tự động, phù hợp ưu tiên người dùng!

---

## 📊 Cấu Trúc File Đầu Vào / Đầu Ra

### Input Files:
```
TKB-20251-K66-69-du-kien-15.07.2025.xlsx  (File gốc - từ người dùng)
```

### Intermediate Files:
```
Ma_hoc_phan_ET_EE.xlsx                (Đã lọc)
Ma_hoc_phan_ET_EE_fixed.xlsx          (Có header chuẩn) ⭐
timetable_all.csv                     (Dataset cho AI)
timetable_user.csv                    (Cấu hình ưu tiên)
classes_to_schedule.csv               (Input cho solver)
timeslots.csv                         (Lưới thời gian)
constraints.json                      (Ràng buộc)
ai_ranked_classes.csv                 (Xếp hạng AI)
```

### Output Files:
```
schedule_final.csv                    (TKB đã xếp tự động) ⭐⭐⭐
```

---

## 🎯 Các Kịch Bản Sử Dụng

### Kịch Bản 1: Chỉ Cần Xếp TKB Tự Động (Nhanh)

```bash
python loc_ma_hoc_phan.py
python build_scheduler_input.py
python greedy_solver.py
```

**Kết quả**: `schedule_final.csv` - TKB hợp lệ nhưng không tối ưu theo sở thích

---

### Kịch Bản 2: AI Gợi Ý + Xếp TKB (Khuyến Nghị)

```bash
python loc_ma_hoc_phan.py
python build_training_dataset.py
python run_pipeline.py
```

**Kết quả**: `schedule_final.csv` - Vừa hợp lệ, vừa phù hợp ưu tiên người dùng

---

### Kịch Bản 3: Chỉ Muốn Xem AI Gợi Ý

```bash
python build_training_dataset.py
python ai_recommender.py
```

**Kết quả**: `ai_ranked_classes.csv` - Xem lớp nào được AI đề xuất cao nhất

---

## ⚙️ Tùy Chỉnh Pipeline

### Thay Đổi Ưu Tiên Người Dùng

Sửa file `timetable_user.csv`:
```csv
PreferredDays,PreferredTimeSlots,AvoidTeachers,PreferredRooms,MaxCredits
Mon,Tue,Wed,Thu,"07:00-11:00,13:00-17:00","Nguyễn Văn A",D5-401,20
```

### Thay Đổi Khung Giờ

Sửa trong `build_scheduler_input.py`:
```python
DEFAULT_SLOTS = [
    {'Slot': 1, 'Start': '07:00', 'End': '09:00'},
    {'Slot': 2, 'Start': '09:00', 'End': '11:00'},
    # Thêm slot mới...
]
```

### Thay Đổi Ràng Buộc

Sửa file `constraints.json`:
```json
{
  "no_overlap": {
    "by": ["Teacher", "RoomAssigned"]
  },
  "max_classes_per_slot": 10
}
```

---

## 🔍 Kiểm Tra Kết Quả

### Xem TKB Đã Xếp:

```bash
# Mở file CSV
notepad schedule_final.csv

# Hoặc dùng Excel
start schedule_final.csv
```

### Kiểm Tra Xung Đột:

- Kiểm tra xem có giáo viên nào bị trùng slot không
- Kiểm tra phòng có bị đặt double booking không
- Xem lớp nào chưa được xếp (Day/TimeSlot/RoomAssigned trống)

---

## ❓ Troubleshooting

### Lỗi: "File không tồn tại"
→ Đảm bảo chạy các bước theo thứ tự, file output của bước trước là input của bước sau

### Lỗi: "0 dòng trong timetable_all.csv"
→ Kiểm tra file Excel có đúng format không, có header chuẩn không

### Lỗi: "AI score = 0 hoặc hằng số"
→ Thu hẹp `PreferredDays/PreferredTimeSlots`, điền `AvoidTeachers/PreferredRooms`, và đảm bảo `timetable_all.csv` có đủ `Thứ/Thời_gian/Phòng`.

### Kết quả: Nhiều lớp chưa được xếp
→ Tăng số slot trong `timeslots.csv` hoặc giảm số lớp trong `classes_to_schedule.csv`

---

**Chúc bạn thành công! 🎉**

