# 📖 Hướng Dẫn Sử Dụng Hệ Thống SPCN Qua Web

> Tài liệu hướng dẫn chi tiết cách sử dụng hệ thống Sắp Xếp Thời Khóa Biểu (SPCN) qua giao diện web

---

## 📋 Mục Lục

1. [Giới Thiệu](#giới-thiệu)
2. [Bắt Đầu Sử Dụng](#bắt-đầu-sử-dụng)
3. [Trang Dashboard](#trang-dashboard)
4. [Trang Upload và Xử Lý Dữ Liệu](#trang-upload-và-xử-lý-dữ-liệu)
5. [Xem Thời Khóa Biểu](#xem-thời-khóa-biểu)
6. [Workflow Hoàn Chỉnh](#workflow-hoàn-chỉnh)
7. [Cấu Hình Ưu Tiên](#cấu-hình-ưu-tiên)
8. [Câu Hỏi Thường Gặp](#câu-hỏi-thường-gặp)

---

## 🎯 Giới Thiệu

Hệ thống **SPCN (Sắp Xếp Thời Khóa Biểu)** là một hệ thống tự động hóa việc sắp xếp thời khóa biểu cho trường học với các tính năng:

- 🤖 **AI Gợi Ý**: Sử dụng AI để gợi ý lớp học phù hợp với sở thích của bạn
- ⚡ **Xếp Lịch Tự Động**: Tự động xếp thời khóa biểu đảm bảo không xung đột
- 🌐 **Giao Diện Web**: Dễ sử dụng, quản lý tập trung qua trình duyệt

### Các Tính Năng Chính

1. **Gợi ý lớp học thông minh**: AI phân tích sở thích của bạn (ngày ưa thích, giờ ưa thích, phòng ưa thích, giáo viên muốn tránh) để gợi ý lớp phù hợp nhất
2. **Xếp lịch tự động**: Tự động xếp thời khóa biểu đảm bảo không xung đột giáo viên/phòng, ưu tiên các lớp được AI đánh giá cao
3. **Xem và quản lý TKB**: Xem thời khóa biểu toàn trường, cá nhân học sinh, hoặc giáo viên

---

## 🚀 Bắt Đầu Sử Dụng

### Bước 1: Truy Cập Hệ Thống

1. Mở trình duyệt web (Chrome, Firefox, Edge, ...)
2. Truy cập địa chỉ: `http://localhost:5000` (hoặc địa chỉ server của bạn)
3. Bạn sẽ thấy trang đăng nhập

### Bước 2: Đăng Nhập / Đăng Ký

#### Nếu chưa có tài khoản:
1. Click nút **"Đăng ký"** hoặc **"Register"**
2. Điền thông tin:
   - **Tên người dùng** (Username)
   - **Email**
   - **Mật khẩu** (Password)
3. Click **"Đăng ký"**
4. Sau khi đăng ký thành công, bạn sẽ được chuyển về trang đăng nhập

#### Nếu đã có tài khoản:
1. Nhập **Tên người dùng** hoặc **Email**
2. Nhập **Mật khẩu**
3. Click **"Đăng nhập"** hoặc **"Login"**
4. Sau khi đăng nhập thành công, bạn sẽ được chuyển đến **Trang Dashboard**

---

## 🏠 Trang Dashboard

Trang Dashboard là trang chính của hệ thống, nơi bạn có thể xem tổng quan và thực hiện các thao tác chính.

### Các Thông Tin Hiển Thị

1. **Thống Kê TKB**:
   - Tổng số lớp học
   - Tổng số phòng học
   - Tổng số môn học
   - Tỉ lệ lấp đầy thời khóa biểu

2. **Lịch Hôm Nay**:
   - Hiển thị các lớp học trong ngày hôm nay
   - Thông tin: Môn học, Giờ, Phòng, Giáo viên

3. **Tổng Quan Tuần**:
   - Lịch học trong tuần (Thứ 2 - Thứ 7)
   - Xem nhanh các lớp học theo từng ngày

4. **Top Gợi Ý AI**:
   - Danh sách các lớp được AI đánh giá cao nhất
   - Hiển thị điểm AI score và thông tin lớp

### Các Nút Thao Tác Nhanh

#### 1. **"Chạy sắp xếp TKB"** 🔄

**Chức năng**: Chạy toàn bộ quy trình xếp lịch tự động (bao gồm AI gợi ý + xếp lịch)

**Cách sử dụng**:
1. Click nút **"Chạy sắp xếp TKB"**
2. Hệ thống sẽ hiển thị cửa sổ tiến trình (Job Status)
3. Bạn có thể xem logs real-time trong cửa sổ này
4. Chờ đến khi thông báo **"Hoàn thành"** hoặc **"Completed"**
5. Kết quả sẽ được lưu vào `schedule_final.csv`

**Lưu ý**:
- Quá trình này có thể mất vài phút tùy vào số lượng lớp
- Bạn có thể đóng cửa sổ tiến trình, nhưng không nên tắt trang web
- Nếu có lỗi, kiểm tra logs để biết nguyên nhân

#### 2. **"Chạy gợi ý lớp"** 🤖

**Chức năng**: Chỉ chạy AI để gợi ý và xếp hạng các lớp học

**Cách sử dụng**:
1. Click nút **"Chạy gợi ý lớp"**
2. Hệ thống sẽ phân tích và đánh giá các lớp dựa trên sở thích của bạn
3. Xem tiến trình trong cửa sổ Job Status
4. Kết quả sẽ được lưu vào `ai_ranked_classes.csv`

**Khi nào sử dụng**:
- Khi bạn muốn xem các lớp được AI gợi ý trước khi xếp lịch
- Khi bạn đã thay đổi cấu hình ưu tiên và muốn xem lại gợi ý

#### 3. **"Tạo TKB gợi ý"** 📅

**Chức năng**: Tạo thời khóa biểu cá nhân từ các lớp được AI gợi ý cao

**Cách sử dụng**:
1. Click nút **"Tạo TKB gợi ý"**
2. Hệ thống sẽ lọc các lớp có điểm AI cao nhất
3. Tạo TKB cá nhân và lưu vào `TKB_ca_nhan.csv` hoặc `schedule_recommended.csv`

**Khi nào sử dụng**:
- Khi bạn muốn xem TKB cá nhân dựa trên gợi ý AI
- Khi bạn chưa muốn chạy xếp lịch tự động đầy đủ

---

## 📤 Trang Upload và Xử Lý Dữ Liệu

Trang Upload cho phép bạn upload file đầu vào và thực hiện các bước xử lý dữ liệu ban đầu.

### Upload File

1. Truy cập trang **Upload** (click vào menu **"Upload"** hoặc truy cập `/upload`)
2. Chọn file cần upload:
   - **File Excel gốc**: `TKB-20251-K66-69-du-kien-15.07.2025.xlsx` (file từ trường)
   - **File CSV**: Các file dữ liệu khác
   - **File JSON**: File cấu hình
3. Click nút **"Upload"** hoặc **"Chọn file"**
4. File sẽ được lưu vào thư mục `data/input/` hoặc `data/output/`

### Các Nút Xử Lý Dữ Liệu

#### 1. **"Lọc mã học phần ET/EE"** 🔍

**Chức năng**: Lọc các mã học phần ET/EE từ file Excel gốc

**Cách sử dụng**:
1. Đảm bảo đã upload file Excel gốc (ví dụ: `TKB-20251-K66-69-du-kien-15.07.2025.xlsx`)
2. Click nút **"Lọc mã học phần ET/EE"**
3. Hệ thống sẽ:
   - Đọc tất cả sheets từ file Excel
   - Tìm các dòng có mã học phần bắt đầu bằng `ET` hoặc `EE`
   - Tạo file mới: `Ma_hoc_phan_ET_EE.xlsx`
   - Tạo file danh sách: `Danh_sach_ma_ET_EE.txt`
4. Xem kết quả trong cửa sổ tiến trình
5. File kết quả sẽ được lưu trong `data/input/` và `data/output/`

**Khi nào sử dụng**:
- Lần đầu tiên xử lý file Excel từ trường
- Khi cần lọc lại dữ liệu từ file mới

#### 2. **"Tạo dataset huấn luyện"** 📊

**Chức năng**: Chuyển đổi dữ liệu Excel thành format CSV chuẩn cho AI

**Cách sử dụng**:
1. Đảm bảo đã có file `Ma_hoc_phan_ET_EE.xlsx` (từ bước trên)
2. Click nút **"Tạo dataset huấn luyện"**
3. Hệ thống sẽ:
   - Đọc file Excel và tự động nhận diện cột (Thứ, Giờ, Phòng, Giáo viên, ...)
   - Chuẩn hóa dữ liệu (chuyển "Thứ 2" → "Mon", "7:00-9:00" → "Slot1", ...)
   - Tạo 2 file CSV:
     - `timetable_all.csv`: Tất cả các lớp (dùng để train AI)
     - `timetable_user.csv`: Cấu hình ưu tiên của bạn
4. File kết quả sẽ được lưu trong `data/output/`

**Khi nào sử dụng**:
- Sau khi đã lọc mã học phần
- Khi cần chuẩn bị dữ liệu cho AI

**Lưu ý**:
- File `timetable_user.csv` chứa cấu hình ưu tiên của bạn, bạn có thể chỉnh sửa file này (xem phần [Cấu Hình Ưu Tiên](#cấu-hình-ưu-tiên))

#### 3. **"Tạo input cho solver"** ⚙️

**Chức năng**: Chuẩn bị dữ liệu đầu vào cho thuật toán xếp lịch

**Cách sử dụng**:
1. Đảm bảo đã có file `Ma_hoc_phan_ET_EE.xlsx` hoặc `timetable_all.csv`
2. Click nút **"Tạo input cho solver"**
3. Hệ thống sẽ:
   - Đọc danh sách lớp từ Excel/CSV
   - Tạo `classes_to_schedule.csv`: Danh sách lớp cần xếp (chưa có Day/TimeSlot/Room)
   - Tạo `timeslots.csv`: Lưới thời gian (24 slots: Thứ 2-7, 4 ca/ngày)
   - Tạo/đọc `constraints.json`: Ràng buộc và ưu tiên
4. File kết quả sẽ được lưu trong `data/output/` và `config/`

**Khi nào sử dụng**:
- Sau khi đã tạo dataset huấn luyện
- Khi cần chuẩn bị dữ liệu cho quá trình xếp lịch

### Preview File

Trang Upload cũng cho phép bạn xem trước (preview) các file kết quả:
- Click vào tên file trong danh sách
- Hệ thống sẽ hiển thị nội dung file (CSV, JSON, hoặc thông tin file Excel)

---

## 📅 Xem Thời Khóa Biểu

Hệ thống cung cấp nhiều cách xem thời khóa biểu khác nhau:

### 1. Xem TKB Toàn Trường (`/timetable/school`)

**Chức năng**: Xem thời khóa biểu của toàn bộ trường

**Cách sử dụng**:
1. Click menu **"Timetable"** → **"TKB Toàn Trường"** hoặc truy cập `/timetable/school`
2. Hệ thống sẽ đọc từ `schedule_final.csv`
3. Hiển thị dạng bảng hoặc lịch tuần với tất cả các lớp

**Thông tin hiển thị**:
- Mã lớp (ClassID)
- Mã học phần (CourseID)
- Tên môn học (SubjectName)
- Giáo viên (Teacher)
- Thứ (Day)
- Ca học (TimeSlot)
- Phòng (Room)

### 2. Xem TKB Cá Nhân Học Sinh (`/timetable/student`)

**Chức năng**: Xem thời khóa biểu cá nhân của học sinh

**Cách sử dụng**:
1. Click menu **"Timetable"** → **"TKB Cá Nhân"** hoặc truy cập `/timetable/student`
2. Hệ thống sẽ đọc từ `TKB_ca_nhan.csv` hoặc `schedule_recommended.csv`
3. Hiển thị lịch học cá nhân của bạn

**Lưu ý**:
- TKB cá nhân được tạo từ các lớp được AI gợi ý
- Bạn có thể xem TKB cá nhân sau khi chạy **"Tạo TKB gợi ý"**

### 3. Xem TKB Cá Nhân Giáo Viên (`/timetable/teacher`)

**Chức năng**: Xem thời khóa biểu của một giáo viên cụ thể

**Cách sử dụng**:
1. Click menu **"Timetable"** → **"TKB Giáo Viên"** hoặc truy cập `/timetable/teacher`
2. Chọn hoặc nhập tên giáo viên
3. Hệ thống sẽ lọc và hiển thị các lớp của giáo viên đó

---

## 🔄 Workflow Hoàn Chỉnh

Dưới đây là quy trình sử dụng hệ thống từ đầu đến cuối:

### Workflow 1: Xếp Lịch Tự Động Hoàn Chỉnh (Khuyến nghị)

```
Bước 1: Đăng nhập vào hệ thống
    ↓
Bước 2: Upload file Excel gốc (Trang Upload)
    ↓
Bước 3: Click "Lọc mã học phần ET/EE"
    ↓
Bước 4: Click "Tạo dataset huấn luyện"
    ↓
Bước 5: (Tùy chọn) Chỉnh sửa timetable_user.csv để cấu hình ưu tiên
    ↓
Bước 6: Click "Tạo input cho solver"
    ↓
Bước 7: Về Trang Dashboard, click "Chạy sắp xếp TKB"
    ↓
Bước 8: Chờ quá trình hoàn thành (xem logs real-time)
    ↓
Bước 9: Xem kết quả tại Trang Timetable Viewer
```

### Workflow 2: Chỉ Xem Gợi Ý AI

```
Bước 1: Đăng nhập vào hệ thống
    ↓
Bước 2: Đảm bảo đã có timetable_all.csv và timetable_user.csv
    ↓
Bước 3: Về Trang Dashboard, click "Chạy gợi ý lớp"
    ↓
Bước 4: Xem kết quả gợi ý trong Dashboard (Top Gợi ý AI)
    ↓
Bước 5: (Tùy chọn) Click "Tạo TKB gợi ý" để xem TKB cá nhân
```

### Workflow 3: Xử Lý File Mới

```
Bước 1: Đăng nhập vào hệ thống
    ↓
Bước 2: Trang Upload → Upload file Excel mới
    ↓
Bước 3: Click "Lọc mã học phần ET/EE"
    ↓
Bước 4: Click "Tạo dataset huấn luyện"
    ↓
Bước 5: Click "Tạo input cho solver"
    ↓
Bước 6: Sẵn sàng để chạy xếp lịch hoặc gợi ý AI
```

---

## ⚙️ Cấu Hình Ưu Tiên

Để AI gợi ý lớp phù hợp với bạn, bạn cần cấu hình ưu tiên trong file `timetable_user.csv`.

### Cách Chỉnh Sửa

1. Truy cập trang **Upload**
2. Tìm file `timetable_user.csv` trong danh sách
3. Click vào file để xem nội dung
4. Download file về máy (nếu cần)
5. Chỉnh sửa file bằng Excel hoặc text editor
6. Upload lại file đã chỉnh sửa

### Các Trường Cấu Hình

File `timetable_user.csv` có định dạng:

```csv
PreferredDays,PreferredTimeSlots,PreferredRooms,AvoidTeachers
Mon|Tue|Wed,Slot1|Slot2,A101|A102|B201,
```

#### 1. **PreferredDays** (Ngày ưa thích)

- Định dạng: Các ngày cách nhau bởi dấu `|`
- Giá trị: `Mon`, `Tue`, `Wed`, `Thu`, `Fri`, `Sat`
- Ví dụ: `Mon|Tue|Wed` (ưa thích Thứ 2, 3, 4)

#### 2. **PreferredTimeSlots** (Giờ ưa thích)

- Định dạng: Các ca cách nhau bởi dấu `|`
- Giá trị: `Slot1`, `Slot2`, `Slot3`, `Slot4`
  - `Slot1`: 7:00-9:00
  - `Slot2`: 9:00-11:00
  - `Slot3`: 13:00-15:00
  - `Slot4`: 15:00-17:00
- Ví dụ: `Slot1|Slot2` (ưa thích ca 1 và ca 2)

#### 3. **PreferredRooms** (Phòng ưa thích)

- Định dạng: Các phòng cách nhau bởi dấu `|`
- Giá trị: Tên phòng (ví dụ: `A101`, `A102`, `B201`, ...)
- Ví dụ: `A101|A102|B201` (ưa thích phòng A101, A102, B201)

#### 4. **AvoidTeachers** (Giáo viên muốn tránh)

- Định dạng: Các tên giáo viên cách nhau bởi dấu `|`
- Giá trị: Tên giáo viên (chính xác như trong dữ liệu)
- Ví dụ: `Nguyễn Văn A|Trần Thị B` (muốn tránh 2 giáo viên này)
- Để trống nếu không muốn tránh ai

### Ví Dụ Cấu Hình Đầy Đủ

```csv
PreferredDays,PreferredTimeSlots,PreferredRooms,AvoidTeachers
Mon|Tue|Wed|Thu,Slot1|Slot2,A101|A102|B201|B202,Nguyễn Văn A
```

**Ý nghĩa**:
- Ưa thích học vào Thứ 2, 3, 4, 5
- Ưa thích học ca 1 (7:00-9:00) và ca 2 (9:00-11:00)
- Ưa thích phòng A101, A102, B201, B202
- Muốn tránh giáo viên "Nguyễn Văn A"

### Sau Khi Chỉnh Sửa

1. Upload lại file `timetable_user.csv` đã chỉnh sửa
2. Chạy lại **"Chạy gợi ý lớp"** hoặc **"Chạy sắp xếp TKB"** để áp dụng cấu hình mới

---

## ❓ Câu Hỏi Thường Gặp

### 1. Làm sao để biết quá trình đang chạy?

- Khi bạn click một nút thao tác, hệ thống sẽ hiển thị cửa sổ **Job Status**
- Cửa sổ này hiển thị logs real-time, bạn có thể xem tiến trình
- Trạng thái: **Running** (đang chạy), **Completed** (hoàn thành), **Failed** (lỗi)

### 2. Quá trình chạy lâu, có thể tắt trình duyệt không?

- **Không nên** tắt trình duyệt khi quá trình đang chạy
- Bạn có thể đóng cửa sổ Job Status, nhưng không nên tắt tab hoặc trình duyệt
- Nếu tắt trình duyệt, quá trình có thể bị gián đoạn

### 3. File kết quả được lưu ở đâu?

- File kết quả được lưu trong:
  - `data/input/`: File đầu vào
  - `data/output/`: File kết quả (CSV, Excel)
  - `config/`: File cấu hình (JSON)
- Bạn có thể xem và download file từ trang Upload

### 4. Làm sao để xem lại kết quả cũ?

- Truy cập trang **Upload**
- Tìm file kết quả trong danh sách (ví dụ: `schedule_final.csv`, `ai_ranked_classes.csv`)
- Click vào file để xem nội dung
- Hoặc xem qua trang **Timetable Viewer**

### 5. Có thể chạy nhiều quá trình cùng lúc không?

- Hệ thống hỗ trợ chạy nhiều job cùng lúc
- Tuy nhiên, để tránh xung đột, nên chờ một quá trình hoàn thành trước khi chạy quá trình khác

### 6. Làm sao để chỉnh sửa ràng buộc xếp lịch?

- File `config/constraints.json` chứa các ràng buộc
- Bạn có thể chỉnh sửa file này (cần kiến thức về JSON)
- Upload lại file sau khi chỉnh sửa

### 7. Hệ thống hỗ trợ những ngành nào?

- Hệ thống hỗ trợ 2 ngành: **EE** (Điện) và **ET** (Điện tử)
- Mỗi ngành có file kết quả riêng (ví dụ: `schedule_final_EE.csv`, `schedule_final_ET.csv`)

### 8. Làm sao để xem TKB của một giáo viên cụ thể?

- Truy cập trang **Timetable Viewer** → **"TKB Giáo Viên"**
- Chọn hoặc nhập tên giáo viên
- Hệ thống sẽ hiển thị các lớp của giáo viên đó

### 9. AI score là gì?

- **AI score** là điểm đánh giá của AI về độ phù hợp của lớp học với sở thích của bạn
- Điểm từ 0 đến 1, điểm càng cao càng phù hợp
- Lớp có điểm cao sẽ được ưu tiên khi xếp lịch

### 10. Làm sao để reset lại toàn bộ dữ liệu?

- Xóa các file trong `data/input/` và `data/output/`
- Upload lại file Excel gốc và chạy lại từ đầu

---

## 🎓 Kết Luận

Hệ thống SPCN cung cấp một giải pháp tự động hóa hoàn chỉnh cho việc sắp xếp thời khóa biểu:

✅ **Dễ sử dụng**: Giao diện web thân thiện, không cần kiến thức lập trình  
✅ **Tự động hóa**: Từ upload file đến xếp lịch, tất cả đều tự động  
✅ **AI thông minh**: Gợi ý lớp phù hợp với sở thích cá nhân  
✅ **Đảm bảo chất lượng**: Không xung đột giáo viên/phòng, ưu tiên lớp phù hợp  

**Quy trình chính**:
```
Upload file → Lọc dữ liệu → Tạo dataset → Cấu hình ưu tiên → 
Chạy AI gợi ý → Xếp lịch tự động → Xem kết quả
```

---

**Tác giả**: Hệ thống SPCN  
**Ngày cập nhật**: 2025-01-09  
**Phiên bản**: 1.0
