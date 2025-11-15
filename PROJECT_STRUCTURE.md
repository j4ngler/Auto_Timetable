# 📁 Cấu Trúc Dự Án SPCN_HaiAnh

## 🗂️ Tổng Quan Cấu Trúc

```
SPCN_HaiAnh/
├── scripts/              # Tất cả các script Python
│   ├── ai_recommender.py
│   ├── build_training_dataset.py
│   ├── build_scheduler_input.py
│   ├── greedy_solver.py
│   ├── loc_ma_hoc_phan.py
│   ├── run_pipeline.py
│   └── ...
│
├── data/                 # Dữ liệu
│   ├── input/           # Dữ liệu đầu vào (Excel gốc)
│   │   ├── TKB-20251-K66-69-du-kien-15.07.2025.xlsx
│   │   ├── Ma_hoc_phan_ET_EE.xlsx
│   │   └── Ma_hoc_phan_ET_EE_fixed.xlsx
│   │
│   └── output/          # Dữ liệu đã xử lý (CSV)
│       ├── timetable_all.csv
│       ├── timetable_user.csv
│       ├── classes_to_schedule.csv
│       ├── ai_ranked_classes.csv
│       ├── schedule_final.csv
│       └── ...
│
├── config/              # File cấu hình
│   └── constraints.json
│
├── docs/                # Tài liệu
│   ├── README.md
│   ├── PIPELINE_HUONG_DAN.md
│   ├── GIAO_AN_DAY_HOC.md
│   └── ...
│
├── web/                 # Ứng dụng web Flask
│   ├── app.py
│   ├── templates/
│   ├── static/
│   └── ...
│
└── [Root files]         # File ở thư mục gốc
    ├── .gitignore
    └── PROJECT_STRUCTURE.md
```

## 📝 Mô Tả Các Thư Mục

### `scripts/`
Chứa tất cả các script Python xử lý dữ liệu, AI, và thuật toán:
- **Xử lý dữ liệu**: `loc_ma_hoc_phan.py`, `build_training_dataset.py`, `build_scheduler_input.py`
- **AI/ML**: `ai_recommender.py`
- **Thuật toán**: `greedy_solver.py`
- **Pipeline**: `run_pipeline.py`
- **Tiện ích**: `recommend_schedule.py`, `random_timetable.py`

**Cách chạy**: Từ thư mục gốc dự án:
```bash
python scripts/ai_recommender.py
python scripts/run_pipeline.py
```

### `data/input/`
Chứa dữ liệu đầu vào gốc (file Excel):
- File Excel từ hệ thống quản lý học phần
- File đã được lọc theo ngành (ET/EE)

### `data/output/`
Chứa tất cả dữ liệu đã xử lý:
- CSV đã chuẩn hóa
- Kết quả AI (ai_ranked_classes.csv)
- Thời khóa biểu đã xếp (schedule_final.csv)
- File cấu hình người dùng (timetable_user.csv)

### `config/`
Chứa file cấu hình hệ thống:
- `constraints.json`: Ràng buộc xếp lịch (không trùng giáo viên/phòng)

### `docs/`
Chứa tất cả tài liệu:
- README.md: Hướng dẫn sử dụng
- PIPELINE_HUONG_DAN.md: Hướng dẫn chi tiết pipeline
- GIAO_AN_DAY_HOC.md: Giáo án giảng dạy
- Slide, PDF, v.v.

### `web/`
Ứng dụng web Flask:
- `app.py`: Backend chính
- `templates/`: HTML templates
- `static/`: CSS, images
- `users.db`: Database người dùng

## 🔄 Tương Thích Ngược

Các script và web app đã được cập nhật để:
1. **Tự động tìm file** ở vị trí mới (`data/output/`, `config/`)
2. **Fallback** về vị trí cũ (thư mục gốc) nếu không tìm thấy
3. **Hoạt động** dù chạy từ thư mục gốc hay từ `scripts/`

## 📌 Lưu Ý

- **Scripts** nên chạy từ **thư mục gốc** dự án để đảm bảo đường dẫn đúng
- **Web app** tự động tìm file ở cả vị trí mới và cũ
- **File mới** sẽ được tạo ở `data/output/` hoặc `config/` tùy loại

## 🚀 Cách Sử Dụng

### Chạy từ thư mục gốc:
```bash
# Chạy pipeline hoàn chỉnh
python scripts/run_pipeline.py

# Chạy AI recommender
python scripts/ai_recommender.py

# Chạy web app
python web/app.py
```


