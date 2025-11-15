# Scripts Directory

Thư mục này chứa tất cả các script Python của dự án.

## Các script chính:

- `loc_ma_hoc_phan.py`: Lọc mã học phần ET/EE từ file Excel gốc
- `build_training_dataset.py`: Tạo dataset huấn luyện từ Excel
- `build_scheduler_input.py`: Tạo input cho solver
- `ai_recommender.py`: Chạy AI để xếp hạng lớp học
- `greedy_solver.py`: Thuật toán xếp thời khóa biểu
- `run_pipeline.py`: Pipeline hoàn chỉnh (AI + Solver)
- `recommend_schedule.py`: Tạo TKB cá nhân

## Cách chạy:

**Lưu ý**: Các script này cần chạy từ thư mục gốc dự án (không phải từ thư mục scripts/).

```bash
# Từ thư mục gốc dự án
python scripts/loc_ma_hoc_phan.py
python scripts/build_training_dataset.py
python scripts/ai_recommender.py
python scripts/run_pipeline.py
```

Hoặc sử dụng các script wrapper trong thư mục gốc.

