# 📚 GIÁO ÁN GIẢNG DẠY: HỆ THỐNG XẾP THỜI KHÓA BIỂU THÔNG MINH

## 🎯 TỔNG QUAN KHÓA HỌC

**Tên dự án**: Smart Timetable System (SPCN_HaiAnh)  
**Tổng thời lượng**: 9-12 giờ (6 buổi, mỗi buổi 1.5-2 giờ)  
**Đối tượng**: Học sinh lớp 12, có kiến thức cơ bản về Python  
**Mục tiêu**: Hiểu cách AI và lập trình được áp dụng vào bài toán thực tế - xếp thời khóa biểu tự động

---

## 📋 PHÂN CHIA CHƯƠNG TRÌNH

### **BUỔI 1: TỔNG QUAN HỆ THỐNG & DEMO**
**Thời lượng**: 2 giờ

#### **Phần 1: Giới thiệu dự án (30 phút)**
- **Mục tiêu học tập**: Hiểu bài toán thực tế và cách AI giải quyết
- **Nội dung** (giải thích đơn giản):
  - **Bài toán**: Xếp thời khóa biểu cho hàng trăm lớp học
  - **Thách thức**: Không được trùng giáo viên, không trùng phòng, phải hợp lý
  - **Giải pháp AI**: Máy tính tự động xếp dựa trên sở thích người dùng
  - **Luồng hoạt động**: Excel → Xử lý dữ liệu → AI gợi ý → Tự động xếp → Web hiển thị
- **Demo trực quan**:
  - Chạy web app, xem giao diện
  - Xem kết quả: thời khóa biểu đã được xếp tự động
  - So sánh: xếp thủ công vs xếp bằng AI

#### **Phần 2: Cài đặt và chạy hệ thống (30 phút)**
- **Mục tiêu**: Biết cách chạy hệ thống có sẵn
- **Nội dung** (hướng dẫn từng bước):
  - Cài đặt Python (nếu chưa có)
  - Giải thích: thư viện là gì? (như các công cụ cần thiết)
  - Cài đặt thư viện: `pip install pandas flask scikit-learn`
  - Cấu trúc thư mục: giải thích các file quan trọng
- **Thực hành**: 
  - Chạy `python web/app.py`
  - Mở trình duyệt: http://127.0.0.1:5000
  - Đăng ký tài khoản, đăng nhập
  - Khám phá các tính năng: Dashboard, TKB, Tạo TKB

#### **Phần 3: Khám phá dữ liệu (60 phút)**
- **Mục tiêu**: Hiểu dữ liệu đầu vào
- **Nội dung** (giải thích dễ hiểu):
  - Mở file Excel gốc: `TKB-20251-K66-69-du-kien-15.07.2025.xlsx`
  - Giải thích các cột: Mã học phần, Tên môn, Giáo viên, Phòng, Thứ, Giờ
  - Vấn đề thực tế: dữ liệu không chuẩn, thiếu thông tin
  - Cách hệ thống xử lý: tự động nhận diện và chuẩn hóa
- **Thực hành**:
  - Mở Excel, quan sát dữ liệu
  - Chạy script xử lý: `python loc_ma_hoc_phan.py`
  - Xem kết quả: file đã được lọc và chuẩn hóa
  - So sánh file trước và sau xử lý

#### **Bài tập về nhà** (đơn giản):
- Đọc `README.md` để hiểu tổng quan
- Thử đăng nhập và khám phá web app

---

### **BUỔI 2: XỬ LÝ DỮ LIỆU VÀ CHUẨN BỊ CHO AI**
**Thời lượng**: 1.5 giờ

#### **Phần 1: Xử lý dữ liệu Excel (45 phút)**
- **Mục tiêu**: Hiểu cách máy tính xử lý dữ liệu không chuẩn
- **Nội dung** (giải thích đơn giản):
  - **Vấn đề**: File Excel có header khác nhau, thiếu dữ liệu
  - **Giải pháp**: Dùng pattern matching (tìm kiếm theo mẫu) để nhận diện cột
  - **Ví dụ**: Tìm cột "Mã học phần" dù có thể ghi là "Mã HP" hoặc "Code"
  - Script `build_training_dataset.py`: tự động chuẩn hóa tất cả
- **Thực hành**:
  - Chạy `python build_training_dataset.py`
  - Xem kết quả: `timetable_all.csv` - dữ liệu đã được chuẩn hóa
  - So sánh: Excel gốc vs CSV đã xử lý

#### **Phần 2: Tạo file cấu hình sở thích (30 phút)**
- **Mục tiêu**: Hiểu cách người dùng thể hiện sở thích
- **Nội dung** (giải thích dễ hiểu):
  - File `timetable_user.csv`: cấu hình sở thích cá nhân
  - **PreferredDays**: Ngày ưa thích (VD: Mon, Wed, Fri)
  - **PreferredTimeSlots**: Khung giờ ưa thích (VD: 07:00-11:00)
  - **PreferredRooms**: Phòng ưa thích (VD: D5-401, C7-205)
  - **AvoidTeachers**: Giáo viên muốn tránh
- **Thực hành**:
  - Mở `timetable_user.csv`, xem cấu trúc
  - Sửa sở thích cá nhân: thích học buổi sáng, thứ 2-4-6
  - Lưu file và quan sát ảnh hưởng

#### **Phần 3: Tạo input cho thuật toán xếp lịch (15 phút)**
- **Mục tiêu**: Hiểu dữ liệu đầu vào cho thuật toán
- **Nội dung** (giải thích ngắn gọn):
  - `classes_to_schedule.csv`: Danh sách lớp cần xếp (chưa có thời gian)
  - `timeslots.csv`: Lưới thời gian (6 ngày × 4 ca = 24 khung giờ)
  - `constraints.json`: Quy tắc (không trùng giáo viên/phòng)
- **Thực hành**:
  - Chạy `python build_scheduler_input.py`
  - Xem các file được tạo
  - Mở `constraints.json`, hiểu các quy tắc

#### **Bài tập về nhà** (đơn giản):
- Tạo `timetable_user.csv` với sở thích của mình
- Thử thay đổi sở thích và quan sát kết quả khác nhau

---

### **BUỔI 3: AI GỢI Ý LỚP HỌC - MACHINE LEARNING**
**Thời lượng**: 2 giờ

#### **Phần 1: AI hoạt động như thế nào? (45 phút)**
- **Mục tiêu**: Hiểu khái niệm cơ bản về AI và Machine Learning
- **Nội dung** (giải thích bằng ví dụ):
  - **AI là gì?**: Máy tính học từ dữ liệu để đưa ra quyết định
  - **Ví dụ đơn giản**: 
    - Nếu bạn thích học thứ 2, 4, 6 → AI sẽ ưu tiên lớp học vào những ngày đó
    - Nếu bạn thích học buổi sáng → AI sẽ ưu tiên lớp học 7h-11h
  - **Cách tính điểm**: 
    - Mỗi lớp được chấm điểm dựa trên sở thích của bạn
    - Lớp nào phù hợp hơn → điểm cao hơn
  - **Random Forest**: Thuật toán AI học từ nhiều quyết định nhỏ
- **Thực hành**:
  - Sửa `timetable_user.csv`: thích học thứ 2-4-6, buổi sáng
  - Chạy `python ai_recommender.py`
  - Xem kết quả: `ai_ranked_classes.csv` - lớp nào được AI đề xuất cao nhất?

#### **Phần 2: Chạy AI và xem kết quả (45 phút)**
- **Mục tiêu**: Thực hành chạy AI và phân tích kết quả
- **Nội dung** (thực hành là chính):
  - Chạy script AI: `python ai_recommender.py`
  - Xem file kết quả: `ai_ranked_classes.csv`
  - Giải thích: Cột `ai_score` - điểm AI càng cao = lớp càng phù hợp
  - So sánh: Lớp có điểm cao vs điểm thấp - tại sao?
- **Thực hành**:
  - Chạy AI với sở thích khác nhau
  - So sánh kết quả: thay đổi PreferredDays → điểm AI thay đổi như thế nào?
  - Mở file CSV, sắp xếp theo điểm AI giảm dần
  - Quan sát: Top 10 lớp được AI đề xuất

#### **Phần 3: Hiểu cách AI học (30 phút)**
- **Mục tiêu**: Hiểu cơ bản về quá trình học của AI
- **Nội dung** (giải thích đơn giản):
  - **Input**: Sở thích của bạn (ngày, giờ, phòng)
  - **Xử lý**: AI so sánh từng lớp với sở thích
  - **Output**: Điểm số cho mỗi lớp
  - **Ví dụ cụ thể**:
    - Lớp học thứ 2, 7h-9h, phòng D5-401
    - Bạn thích: thứ 2 (+1 điểm), 7h-11h (+1 điểm), phòng D5 (+1 điểm)
    - Tổng: 3 điểm
  - **Random Forest**: Dùng nhiều "cây quyết định" để dự đoán chính xác hơn
- **Thực hành**:
  - Thử nghiệm: Sửa sở thích rất cụ thể (chỉ thích 1 ngày, 1 khung giờ)
  - Chạy lại AI → quan sát điểm AI phân hóa rõ hơn
  - Giải thích: Sở thích càng cụ thể → AI càng chính xác

#### **Bài tập về nhà** (đơn giản):
- Thử nghiệm với 3 bộ sở thích khác nhau
- Ghi lại top 5 lớp được AI đề xuất cho mỗi bộ sở thích
- So sánh và giải thích sự khác biệt

---

### **BUỔI 4: TỰ ĐỘNG XẾP THỜI KHÓA BIỂU**
**Thời lượng**: 1.5 giờ

#### **Phần 1: Bài toán xếp lịch (30 phút)**
- **Mục tiêu**: Hiểu các quy tắc khi xếp thời khóa biểu
- **Nội dung** (giải thích bằng ví dụ):
  - **Quy tắc cứng** (bắt buộc):
    - Không được trùng giáo viên: 1 giáo viên không thể dạy 2 lớp cùng lúc
    - Không được trùng phòng: 1 phòng không thể có 2 lớp cùng lúc
  - **Quy tắc mềm** (ưu tiên):
    - Ưu tiên xếp lớp có điểm AI cao trước
    - Ưu tiên xếp vào khung giờ người dùng thích
  - **Ví dụ cụ thể**:
    - Lớp A: Giáo viên X, cần phòng D5-401
    - Lớp B: Giáo viên X, cần phòng D5-401
    - → Không thể xếp cùng 1 khung giờ!
- **Thực hành**:
  - Mở `constraints.json`, xem các quy tắc
  - Giải thích: Tại sao cần các quy tắc này?

#### **Phần 2: Thuật toán Greedy - Xếp lịch tự động (45 phút)**
- **Mục tiêu**: Hiểu cách máy tính tự động xếp lịch
- **Nội dung** (giải thích đơn giản):
  - **Greedy = Tham lam**: Chọn giải pháp tốt nhất ngay tại mỗi bước
  - **Cách hoạt động**:
    1. Lấy lớp có điểm AI cao nhất
    2. Tìm khung giờ đầu tiên còn trống (không trùng giáo viên/phòng)
    3. Gán lớp vào khung giờ đó
    4. Lặp lại với lớp tiếp theo
  - **Ví dụ minh họa**:
    - Lớp 1 (điểm 5.0) → Tìm slot trống → Gán vào Thứ 2, Ca 1
    - Lớp 2 (điểm 4.5) → Tìm slot trống → Gán vào Thứ 2, Ca 2
    - Lớp 3 (điểm 4.0) → Giáo viên trùng với Lớp 1 → Bỏ qua slot đó, tìm slot khác
- **Thực hành**:
  - Chạy `python greedy_solver.py`
  - Xem kết quả: `schedule_final.csv`
  - Kiểm tra: Có lớp nào bị trùng giáo viên/phòng không?
  - Đếm: Có bao nhiêu lớp đã được xếp? Bao nhiêu lớp chưa xếp?

#### **Phần 3: Pipeline hoàn chỉnh - AI + Xếp lịch (15 phút)**
- **Mục tiêu**: Chạy toàn bộ quy trình từ đầu đến cuối
- **Nội dung** (thực hành):
  - Script `run_pipeline.py`: Tự động chạy AI → Xếp lịch
  - Quy trình:
    1. Chạy AI để xếp hạng lớp (nếu chưa có)
    2. Sắp xếp lớp theo điểm AI (cao → thấp)
    3. Chạy thuật toán xếp lịch
    4. Xuất kết quả: `schedule_final.csv`
- **Thực hành**:
  - Chạy `python run_pipeline.py`
  - Quan sát quá trình: AI chạy → Xếp lịch
  - Xem kết quả cuối cùng: Thời khóa biểu đã được xếp tự động
  - So sánh với TKB thủ công: Nhanh hơn, chính xác hơn?

#### **Bài tập về nhà** (đơn giản):
- Chạy pipeline với sở thích khác nhau
- So sánh kết quả: Số lớp được xếp, phân bố theo ngày
- Viết nhận xét: Thuật toán xếp lịch có hợp lý không?

---

### **BUỔI 5: ỨNG DỤNG WEB - GIAO DIỆN NGƯỜI DÙNG**
**Thời lượng**: 2 giờ

#### **Phần 1: Khám phá Web App (30 phút)**
- **Mục tiêu**: Hiểu các tính năng của web app
- **Nội dung** (thực hành):
  - Chạy web app: `python web/app.py`
  - Đăng ký/Đăng nhập: Tạo tài khoản, chọn ngành (EE/ET)
  - Dashboard: Xem thống kê, lịch hôm nay, top gợi ý AI
  - TKB cá nhân: Xem thời khóa biểu của mình
  - TKB toàn trường: Xem tất cả lớp đã được xếp
  - Tạo TKB: Chạy AI và xếp lịch từ giao diện web
- **Thực hành**:
  - Đăng ký tài khoản mới
  - Khám phá từng trang, hiểu chức năng
  - Thử chạy "Chạy gợi ý lớp" và "Chạy sắp xếp TKB"

#### **Phần 2: Cách Web App hoạt động (45 phút)**
- **Mục tiêu**: Hiểu cơ bản về web application
- **Nội dung** (giải thích đơn giản):
  - **Frontend**: Giao diện người dùng (HTML, CSS)
  - **Backend**: Xử lý logic, chạy AI, xếp lịch (Python Flask)
  - **Database**: Lưu thông tin người dùng (SQLite)
  - **Luồng hoạt động**:
    1. Người dùng bấm nút "Chạy gợi ý lớp"
    2. Web app chạy script AI ở phía sau
    3. Hiển thị kết quả trên màn hình
  - **Job System**: Khi chạy AI/xếp lịch (mất thời gian) → tạo "job" để theo dõi
- **Thực hành**:
  - Bấm "Chạy gợi ý lớp" → quan sát job đang chạy
  - Xem log real-time: AI đang làm gì?
  - Chờ kết quả → xem top gợi ý trên Dashboard

#### **Phần 3: Tùy chỉnh và sử dụng (45 phút)**
- **Mục tiêu**: Biết cách sử dụng hệ thống hiệu quả
- **Nội dung** (thực hành):
  - **Upload file**: Tải lên `timetable_user.csv` với sở thích mới
  - **Constraints**: Chỉnh sửa quy tắc xếp lịch (nếu cần)
  - **Profile**: Cập nhật thông tin, đổi ngành
  - **Download**: Tải TKB về máy (CSV)
  - **Preview**: Xem trước file CSV trước khi tải
- **Thực hành**:
  - Upload `timetable_user.csv` với sở thích của mình
  - Chạy lại AI → xem kết quả thay đổi
  - Tải TKB về, mở bằng Excel
  - So sánh TKB cá nhân vs TKB toàn trường

#### **Bài tập về nhà** (đơn giản):
- Tạo TKB cá nhân với sở thích của mình
- Tải về và trình bày: TKB này có hợp lý không?
- Viết nhận xét: Hệ thống có hữu ích không? Cần cải thiện gì?

---

### **BUỔI 6: TỔNG KẾT & DEMO DỰ ÁN**
**Thời lượng**: 2 giờ

#### **Phần 1: Tổng kết kiến thức (30 phút)**
- **Mục tiêu**: Ôn lại toàn bộ quy trình
- **Nội dung**:
  - Review lại 5 buổi học:
    1. Tổng quan hệ thống
    2. Xử lý dữ liệu
    3. AI gợi ý lớp học
    4. Tự động xếp thời khóa biểu
    5. Ứng dụng web
  - **Luồng hoàn chỉnh**: Excel → Xử lý → AI → Xếp lịch → Web
  - **Điểm mạnh của hệ thống**: Tự động, nhanh, chính xác
  - **Ứng dụng thực tế**: Có thể dùng cho trường học, công ty

#### **Phần 2: Demo dự án (60 phút)**
- **Mục tiêu**: Trình bày dự án hoàn chỉnh
- **Nội dung** (thực hành):
  - **Chuẩn bị demo**:
    - Tạo tài khoản mới
    - Cấu hình sở thích cá nhân
    - Chạy toàn bộ pipeline
  - **Trình bày**:
    - Giới thiệu bài toán: Tại sao cần xếp TKB tự động?
    - Demo từng bước:
      1. Upload dữ liệu Excel
      2. Chạy AI gợi ý
      3. Xếp lịch tự động
      4. Xem kết quả trên web
    - So sánh: Xếp thủ công vs Xếp bằng AI
  - **Thảo luận**:
    - Hệ thống có hữu ích không?
    - Cần cải thiện gì?
    - Có thể áp dụng ở đâu?

#### **Phần 3: Q&A và Đánh giá (30 phút)**
- **Mục tiêu**: Giải đáp thắc mắc và nhận phản hồi
- **Nội dung**:
  - Học sinh đặt câu hỏi về:
    - Cách AI hoạt động
    - Thuật toán xếp lịch
    - Cách sử dụng web app
  - Giáo viên đánh giá:
    - Mức độ hiểu bài
    - Khả năng thực hành
    - Ý tưởng cải thiện
  - **Bài tập cuối khóa** (tùy chọn):
    - Viết báo cáo ngắn (1-2 trang) về dự án
    - Trình bày ý tưởng cải thiện hệ thống
    - Demo cho bạn bè/người thân

---

## 📊 BẢNG PHÂN BỔ THỜI LƯỢNG

| Buổi | Nội dung chính | Thời lượng | Tỷ lệ |
|------|----------------|------------|-------|
| 1 | Tổng quan & Demo | 2h | 17% |
| 2 | Xử lý dữ liệu | 1.5h | 13% |
| 3 | AI gợi ý lớp học | 2h | 17% |
| 4 | Tự động xếp TKB | 1.5h | 13% |
| 5 | Ứng dụng Web | 2h | 17% |
| 6 | Tổng kết & Demo | 2h | 17% |

**Tổng**: 11 giờ (6 buổi, mỗi buổi 1.5-2 giờ)

---

## 🎯 MỤC TIÊU HỌC TẬP TỔNG THỂ

Sau khóa học, học sinh lớp 12 có thể:
1. ✅ **Hiểu bài toán thực tế**: Xếp thời khóa biểu là gì? Tại sao cần tự động hóa?
2. ✅ **Hiểu cách AI hoạt động**: Máy tính học từ sở thích người dùng như thế nào?
3. ✅ **Hiểu thuật toán cơ bản**: Greedy algorithm - chọn giải pháp tốt nhất tại mỗi bước
4. ✅ **Sử dụng hệ thống**: Biết cách chạy web app, cấu hình sở thích, xem kết quả
5. ✅ **Đánh giá kết quả**: Phân tích TKB được tạo, so sánh với xếp thủ công
6. ✅ **Ứng dụng thực tế**: Hiểu cách AI và lập trình giải quyết vấn đề thực tế

---

## 📝 TÀI LIỆU THAM KHẢO

- **Code**: Toàn bộ file trong repo `SPCN_HaiAnh`
- **Documentation**: `README.md`, `PIPELINE_HUONG_DAN.md`
- **Video hướng dẫn** (nếu có): Demo từng bước chạy hệ thống
- **Tài liệu Python cơ bản**: Cho học sinh chưa quen với Python

---

## 💡 GỢI Ý ĐIỀU CHỈNH

### **Cho lớp có ít thời gian (6-8 giờ)**:
- Rút gọn buổi 2, 4 (tập trung vào demo)
- Bỏ phần thực hành phức tạp, chỉ giữ demo
- Tăng thời gian buổi 1 (tổng quan) và buổi 6 (tổng kết)

### **Cho lớp có nhiều thời gian (15-18 giờ)**:
- Thêm buổi về Python cơ bản (nếu học sinh chưa biết)
- Thêm buổi thực hành: Tự tạo sở thích và chạy pipeline
- Thêm buổi về cải thiện: Gợi ý cách nâng cấp hệ thống

### **Cho lớp chuyên Tin học**:
- Tăng thời lượng buổi 3 (AI) lên 3 giờ
- Thêm phần giải thích code: Đọc và hiểu script Python
- Thêm bài tập: Sửa code để thay đổi cách tính điểm AI

### **Lưu ý khi dạy học sinh lớp 12**:
- ✅ **Giải thích đơn giản**: Tránh thuật ngữ kỹ thuật, dùng ví dụ cụ thể
- ✅ **Tập trung vào demo**: Học sinh thấy kết quả → hiểu cách hoạt động
- ✅ **Không yêu cầu code**: Chủ yếu chạy script có sẵn, không cần viết code mới
- ✅ **Khuyến khích thử nghiệm**: Thay đổi sở thích, quan sát kết quả
- ✅ **Kết nối thực tế**: Giải thích ứng dụng AI trong đời sống

---

**Chúc bạn soạn bài giảng thành công! 🎓**

