import os
import sys
import threading
import uuid
import subprocess
import json
import sqlite3
from datetime import datetime
from typing import Dict, Any, Optional
import webbrowser

from flask import Flask, render_template, request, send_file, jsonify, abort, redirect, url_for, session
from werkzeug.security import generate_password_hash, check_password_hash


def create_app() -> Flask:
    app = Flask(__name__)
    app.secret_key = os.environ.get('SPCN_SECRET', 'dev-secret-key-change')

    # Project base (parent of web/)
    app.config['PROJECT_BASE_DIR'] = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))

    # Simple in-memory job store
    app.job_store: Dict[str, Dict[str, Any]] = {}
    app.users_file = os.path.join(os.path.dirname(__file__), 'users.json')
    app.users_db = os.path.join(os.path.dirname(__file__), 'users.db')

    # Whitelist files that can be previewed/downloaded/uploaded (add as needed)
    app.allowed_files = {
        'schedule_recommended.csv',
        'schedule_final.csv',
        'classes_to_schedule.csv',
        'timetable_all.csv',
        'ai_ranked_classes.csv',
        'timeslots.csv',
        'constraints.json',
        'timetable_user.csv',
        'TKB_ca_nhan.csv',
    }

    def project_path(filename: str) -> str:
        """Tìm file ở vị trí mới (data/output/, config/) hoặc vị trí cũ (thư mục gốc)"""
        # Đường dẫn mới
        data_output = os.path.join(app.config['PROJECT_BASE_DIR'], 'data', 'output')
        config_dir = os.path.join(app.config['PROJECT_BASE_DIR'], 'config')
        
        # File config
        if filename in ('constraints.json',):
            new_path = os.path.join(config_dir, filename)
            if os.path.exists(new_path):
                return new_path
        
        # File output
        if filename.endswith('.csv') or filename.endswith('.txt'):
            new_path = os.path.join(data_output, filename)
            if os.path.exists(new_path):
                return new_path
        
        # Fallback: tìm ở thư mục gốc (tương thích ngược)
        return os.path.join(app.config['PROJECT_BASE_DIR'], filename)

    def resolve_major_path(filename: str) -> str:
        """Nếu có phiên bản theo major (EE/ET) thì ưu tiên dùng.
        Ví dụ: schedule_final.csv -> schedule_final_EE.csv
        """
        major = session.get('user_major')
        if major in ('EE','ET') and filename in ('schedule_final.csv', 'ai_ranked_classes.csv'):
            base, ext = os.path.splitext(filename)
            candidate = f"{base}_{major}{ext}"
            cand_abs = project_path(candidate)
            if os.path.exists(cand_abs):
                return cand_abs
        return project_path(filename)

    # Warm cache theo major (tạo file _EE/_ET nếu chưa có) chạy nền
    def warm_cache_for_major(major: Optional[str]):
        if major not in ('EE','ET'):
            return
        # Nếu đã có file theo major thì bỏ qua
        base = os.path.join(app.config['PROJECT_BASE_DIR'], f'schedule_final_{major}.csv')
        ai_base = os.path.join(app.config['PROJECT_BASE_DIR'], f'ai_ranked_classes_{major}.csv')
        need_ai = not os.path.exists(ai_base)
        need_sched = not os.path.exists(base)
        # Nếu thiếu AI → chạy gợi ý; nếu thiếu schedule → chạy pipeline
        try:
            if need_ai:
                start_job(f'Warming AI cache ({major})', job_run_recommender, args=(['--major', major],))
            if need_sched:
                start_job(f'Warming schedule cache ({major})', job_run_scheduler, args=(['--major', major],))
        except Exception:
            pass

    def start_job(title: str, target, args=()):
        job_id = str(uuid.uuid4())
        app.job_store[job_id] = {
            'id': job_id,
            'title': title,
            'status': 'running',
            'created_at': datetime.utcnow().isoformat() + 'Z',
            'logs': [],
            'result': None,
            'error': None,
        }

        def _run():
            try:
                result = target(job_id, *args)
                app.job_store[job_id]['status'] = 'completed'
                app.job_store[job_id]['result'] = result
            except Exception as exc:  # noqa: BLE001 - surface any runtime issue to UI
                app.job_store[job_id]['status'] = 'failed'
                app.job_store[job_id]['error'] = str(exc)

        t = threading.Thread(target=_run, daemon=True)
        t.start()
        return job_id

    def run_script(job_id: str, script: str, args=None):
        if args is None:
            args = []
        # Tìm script ở thư mục scripts/ hoặc thư mục gốc
        script_in_scripts = os.path.join(app.config['PROJECT_BASE_DIR'], 'scripts', script)
        script_in_root = project_path(script)
        
        if os.path.exists(script_in_scripts):
            script_abs = script_in_scripts
        elif os.path.exists(script_in_root):
            script_abs = script_in_root
        else:
            raise FileNotFoundError(f"Không tìm thấy script: {script}")

        cmd = [sys.executable, script_abs] + args
        # Đảm bảo chạy từ thư mục gốc dự án
        proc = subprocess.Popen(
            cmd,
            cwd=app.config['PROJECT_BASE_DIR'],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            env=dict(os.environ, PYTHONIOENCODING='utf-8'),  # Đảm bảo encoding UTF-8
        )
        logs = []
        for line in proc.stdout:  # type: ignore[union-attr]
            line = line.rstrip('\n')
            logs.append(line)
            app.job_store[job_id]['logs'] = logs[-500:]
        proc.wait()
        if proc.returncode != 0:
            raise RuntimeError(f"Script lỗi mã {proc.returncode}")
        return {'returncode': proc.returncode}

    # --- SQLite user storage ---
    def _db_conn():
        conn = sqlite3.connect(app.users_db)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db():
        conn = _db_conn()
        try:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS users (
                  email TEXT PRIMARY KEY,
                  username TEXT,
                  name TEXT,
                  mode TEXT,
                  major TEXT,
                  phone TEXT,
                  cohort TEXT,
                  class_name TEXT,
                  avatar_url TEXT,
                  password_hash TEXT
                )
                """
            )
            conn.commit()
            # Bảo đảm các cột mới tồn tại khi nâng cấp từ bản cũ
            def ensure_col(col: str):
                cur = conn.execute("PRAGMA table_info(users)")
                cols = [r[1] for r in cur.fetchall()]
                if col not in cols:
                    conn.execute(f"ALTER TABLE users ADD COLUMN {col} TEXT")
                    conn.commit()
            for c in ('phone','cohort','class_name','avatar_url'):
                ensure_col(c)
        finally:
            conn.close()

    def _get_user_by_email(email: str) -> Optional[Dict[str, Any]]:
        conn = _db_conn()
        try:
            cur = conn.execute('SELECT * FROM users WHERE email = ?', (email,))
            row = cur.fetchone()
            return dict(row) if row else None
        finally:
            conn.close()

    def _get_user_by_username(username: str) -> Optional[Dict[str, Any]]:
        conn = _db_conn()
        try:
            cur = conn.execute('SELECT * FROM users WHERE lower(username) = lower(?)', (username,))
            row = cur.fetchone()
            return dict(row) if row else None
        finally:
            conn.close()

    def _create_user(email: str, username: str, name: str, mode: str, major: str, password_hash: str) -> None:
        conn = _db_conn()
        try:
            conn.execute(
                'INSERT INTO users(email, username, name, mode, major, password_hash) VALUES (?,?,?,?,?,?)',
                (email, username, name, mode, major, password_hash)
            )
            conn.commit()
        finally:
            conn.close()

    def _update_user_major(email: str, major: str) -> None:
        conn = _db_conn()
        try:
            conn.execute('UPDATE users SET major = ? WHERE email = ?', (major, email))
            conn.commit()
        finally:
            conn.close()

    def _update_user_profile(email: str, name: str, phone: str, cohort: str, class_name: str, major: str, avatar_url: str) -> None:
        conn = _db_conn()
        try:
            conn.execute(
                'UPDATE users SET name=?, phone=?, cohort=?, class_name=?, major=?, avatar_url=? WHERE email=?',
                (name, phone, cohort, class_name, major, avatar_url, email)
            )
            conn.commit()
        finally:
            conn.close()

    def _migrate_users_json_to_db() -> None:
        # If db empty and users.json exists, import
        if not os.path.exists(app.users_file):
            return
        conn = _db_conn()
        try:
            cur = conn.execute('SELECT COUNT(1) as c FROM users')
            count = cur.fetchone()[0]
        finally:
            conn.close()
        if count > 0:
            return
        try:
            with open(app.users_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception:
            data = {}
        for email, u in data.items():
            _create_user(
                email=email,
                username=u.get('username') or '',
                name=u.get('name') or email,
                mode=u.get('mode') or 'student',
                major=u.get('major') or None,
                password_hash=u.get('password_hash') or ''
            )

    # Job targets
    def job_run_scheduler(job_id: str, extra_args=None):
        # Ưu tiên chạy run_pipeline.py nếu có; fallback greedy_solver.py
        scripts_dir = os.path.join(app.config['PROJECT_BASE_DIR'], 'scripts')
        if os.path.exists(os.path.join(scripts_dir, 'run_pipeline.py')):
            return run_script(job_id, 'run_pipeline.py', args=(extra_args or []))
        elif os.path.exists(os.path.join(scripts_dir, 'greedy_solver.py')):
            return run_script(job_id, 'greedy_solver.py', args=(extra_args or []))
        else:
            raise FileNotFoundError('Không tìm thấy run_pipeline.py hoặc greedy_solver.py')

    def job_run_recommender(job_id: str, extra_args=None):
        # Gợi ý lớp học kỳ tiếp theo: chạy ai_recommender.py (có thể kèm --major)
        scripts_dir = os.path.join(app.config['PROJECT_BASE_DIR'], 'scripts')
        if os.path.exists(os.path.join(scripts_dir, 'ai_recommender.py')):
            return run_script(job_id, 'ai_recommender.py', args=(extra_args or []))
        else:
            raise FileNotFoundError('Không tìm thấy ai_recommender.py')

    def job_run_recommend_schedule(job_id: str):
        # Tạo TKB gợi ý cá nhân từ ai_ranked_classes.csv
        scripts_dir = os.path.join(app.config['PROJECT_BASE_DIR'], 'scripts')
        if os.path.exists(os.path.join(scripts_dir, 'recommend_schedule.py')):
            return run_script(job_id, 'recommend_schedule.py')
        else:
            raise FileNotFoundError('Không tìm thấy recommend_schedule.py')

    def job_run_loc_ma_hoc_phan(job_id: str):
        # Lọc mã học phần ET/EE từ file Excel gốc
        scripts_dir = os.path.join(app.config['PROJECT_BASE_DIR'], 'scripts')
        if os.path.exists(os.path.join(scripts_dir, 'loc_ma_hoc_phan.py')):
            return run_script(job_id, 'loc_ma_hoc_phan.py')
        else:
            raise FileNotFoundError('Không tìm thấy loc_ma_hoc_phan.py')

    def job_run_build_training_dataset(job_id: str):
        # Tạo dataset huấn luyện từ Excel
        scripts_dir = os.path.join(app.config['PROJECT_BASE_DIR'], 'scripts')
        if os.path.exists(os.path.join(scripts_dir, 'build_training_dataset.py')):
            return run_script(job_id, 'build_training_dataset.py')
        else:
            raise FileNotFoundError('Không tìm thấy build_training_dataset.py')

    def job_run_build_scheduler_input(job_id: str):
        # Tạo input cho solver
        scripts_dir = os.path.join(app.config['PROJECT_BASE_DIR'], 'scripts')
        if os.path.exists(os.path.join(scripts_dir, 'build_scheduler_input.py')):
            return run_script(job_id, 'build_scheduler_input.py')
        else:
            raise FileNotFoundError('Không tìm thấy build_scheduler_input.py')

    @app.route('/')
    def index():
        # Nếu chưa đăng nhập, hiển thị trang chào mừng công khai (không có sidebar/dashboard)
        if not session.get('user_email'):
            return render_template('index_public.html')
        return render_template('index.html')

    # Marketing/static-like pages
    @app.route('/landing')
    def landing():
        # Landing đã được loại bỏ → điều hướng về Dashboard
        return redirect(url_for('index'))

    @app.route('/auth/signin')
    def auth_signin():
        return render_template('auth_signin.html', message=request.args.get('m'), error=request.args.get('e'))

    _init_db()
    _migrate_users_json_to_db()

    @app.route('/auth/login')
    def auth_login():
        return render_template('auth_login.html', message=request.args.get('m'), error=request.args.get('e'))

    @app.route('/auth/logout')
    def auth_logout():
        session.clear()
        return redirect(url_for('index'))

    @app.post('/auth/signin')
    def auth_signin_post():
        email = (request.form.get('email') or '').strip().lower()
        username = (request.form.get('username') or '').strip()
        password = request.form.get('password') or ''
        password2 = request.form.get('password2') or ''
        user_mode = (request.form.get('user_mode') or 'student').strip()
        user_major = (request.form.get('user_major') or 'EE').strip()
        if not email or not password:
            return redirect(url_for('auth_signin', e='Thiếu email hoặc mật khẩu'))
        if password != password2:
            return redirect(url_for('auth_signin', e='Mật khẩu nhập lại không khớp'))
        # check exists
        if _get_user_by_email(email):
            return redirect(url_for('auth_signin', e='Email đã tồn tại'))
        _create_user(
            email=email,
            username=username,
            name=username or email,
            mode=user_mode,
            major=user_major,
            password_hash=generate_password_hash(password)
        )
        return redirect(url_for('auth_login', m='Tạo tài khoản thành công. Vui lòng đăng nhập.'))

    @app.post('/auth/login')
    def auth_login_post():
        email_or_username = (request.form.get('email') or '').strip().lower()
        password = request.form.get('password') or ''
        u = _get_user_by_email(email_or_username)
        if not u:
            u = _get_user_by_username(email_or_username)
        if not u or not check_password_hash(u.get('password_hash', ''), password):
            return redirect(url_for('auth_login', e='Sai email hoặc mật khẩu'))
        # Lưu session
        session['user_email'] = u.get('email') or email_or_username
        session['user_name'] = u.get('name') or u.get('email') or email_or_username
        session['user_major'] = u.get('major')
        # Khởi động làm nóng cache theo major (nếu thiếu) để lần xem sau không phải chạy tay
        warm_cache_for_major(session.get('user_major'))
        return redirect(url_for('index'))

    @app.route('/about')
    def about():
        return render_template('about.html')

    @app.route('/help')
    def help():
        return render_template('help.html')

    @app.route('/contact')
    def contact():
        return render_template('contact.html')

    @app.route('/profile', methods=['GET', 'POST'])
    def profile():
        if not session.get('user_email'):
            return redirect(url_for('auth_login', e='Vui lòng đăng nhập'))
        if request.method == 'GET':
            # Đọc hồ sơ hiện tại từ DB để hiển thị
            u = _get_user_by_email(session.get('user_email')) or {}
            return render_template(
                'profile.html',
                user_email=session.get('user_email'),
                user_name=u.get('name') or session.get('user_name'),
                user_major=u.get('major') or 'EE',
                user_phone=u.get('phone') or '',
                user_cohort=u.get('cohort') or '',
                user_class=u.get('class_name') or '',
                user_avatar=u.get('avatar_url') or ''
            )
        # POST → cập nhật major
        major = (request.form.get('major') or '').strip()
        name = (request.form.get('name') or '').strip()
        phone = (request.form.get('phone') or '').strip()
        cohort = (request.form.get('cohort') or '').strip()
        class_name = (request.form.get('class_name') or '').strip()
        avatar_url = (request.form.get('avatar_url') or '').strip()
        if major not in ('EE','ET'):
            major = session.get('user_major') or 'EE'
        _update_user_profile(session.get('user_email'), name, phone, cohort, class_name, major, avatar_url)
        session['user_major'] = major
        # Làm nóng cache cho major mới chọn (nếu thiếu)
        warm_cache_for_major(major)
        if name:
            session['user_name'] = name
        return redirect(url_for('profile'))

    @app.route('/run/schedule', methods=['POST'])
    def run_schedule():
        major = session.get('user_major')
        if major in ('EE','ET'):
            job_id = start_job('Chạy sắp xếp thời khoá biểu', job_run_scheduler, args=(['--major', major],))
        else:
            job_id = start_job('Chạy sắp xếp thời khoá biểu', job_run_scheduler)
        return jsonify({'job_id': job_id})

    @app.route('/run/recommend', methods=['POST'])
    def run_recommend():
        major = request.args.get('major') or session.get('user_major')
        if major in ('EE','ET'):
            job_id = start_job('Chạy gợi ý lớp học kỳ tiếp theo', job_run_recommender, args=(['--major', major],))
        else:
            job_id = start_job('Chạy gợi ý lớp học kỳ tiếp theo', job_run_recommender)
        return jsonify({'job_id': job_id})

    @app.route('/run/recommend_schedule', methods=['POST'])
    def run_recommend_schedule():
        job_id = start_job('Tạo TKB gợi ý cá nhân', job_run_recommend_schedule)
        return jsonify({'job_id': job_id})

    @app.route('/run/loc_ma_hoc_phan', methods=['POST'])
    def run_loc_ma_hoc_phan():
        job_id = start_job('Lọc mã học phần ET/EE', job_run_loc_ma_hoc_phan)
        return jsonify({'job_id': job_id})

    @app.route('/run/build_training_dataset', methods=['POST'])
    def run_build_training_dataset():
        job_id = start_job('Tạo dataset huấn luyện', job_run_build_training_dataset)
        return jsonify({'job_id': job_id})

    @app.route('/run/build_scheduler_input', methods=['POST'])
    def run_build_scheduler_input():
        job_id = start_job('Tạo input cho solver', job_run_build_scheduler_input)
        return jsonify({'job_id': job_id})

    @app.route('/status/<job_id>')
    def status(job_id: str):
        job = app.job_store.get(job_id)
        if not job:
            abort(404)
        return jsonify(job)

    @app.route('/api/stats')
    def api_stats():
        """Trả về thống kê từ các file CSV để hiển thị dashboard"""
        stats = {
            'total_classes': 0,
            'total_rooms': 0,
            'total_courses': 0,
            'classes_by_day': {'Mon': 0, 'Tue': 0, 'Wed': 0, 'Thu': 0, 'Fri': 0, 'Sat': 0},
            # Ma trận Ngày × Ca (slot) để vẽ heatmap
            'classes_by_day_slot': {
                'Mon': {'1': 0, '2': 0, '3': 0, '4': 0},
                'Tue': {'1': 0, '2': 0, '3': 0, '4': 0},
                'Wed': {'1': 0, '2': 0, '3': 0, '4': 0},
                'Thu': {'1': 0, '2': 0, '3': 0, '4': 0},
                'Fri': {'1': 0, '2': 0, '3': 0, '4': 0},
                'Sat': {'1': 0, '2': 0, '3': 0, '4': 0},
            },
            'top_recommendations': [],
            'total_slots': 24,  # 6 ngày × 4 ca
            'filled_slots': 0,
        }
        
        try:
            import csv
            schedule_path = resolve_major_path('schedule_final.csv')
            if os.path.exists(schedule_path):
                rooms_set = set()
                courses_set = set()
                user_major = session.get('user_major')
                with open(schedule_path, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        # Nếu đang xem theo major nhưng file không tách theo major, lọc theo prefix CourseID
                        if user_major in ('EE','ET'):
                            cid = (row.get('CourseID') or '').strip()
                            if cid and not cid.startswith(user_major):
                                continue
                        stats['total_classes'] += 1
                        day = (row.get('Day') or '').strip()
                        slot = (row.get('TimeSlot') or '').strip()
                        if day in stats['classes_by_day']:
                            stats['classes_by_day'][day] += 1
                            # Tăng theo slot nếu hợp lệ (1-4)
                            if slot in ('1', '2', '3', '4'):
                                stats['classes_by_day_slot'][day][slot] += 1
                        room = (row.get('RoomAssigned') or '').strip()
                        if room:
                            rooms_set.add(room)
                        course = (row.get('CourseID') or '').strip()
                        if course:
                            courses_set.add(course)
                stats['total_rooms'] = len(rooms_set)
                stats['total_courses'] = len(courses_set)
                
                # Tính filled_slots (số slot đã có ít nhất 1 lớp)
                filled_slots_set = set()
                with open(schedule_path, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        if user_major in ('EE','ET'):
                            cid = (row.get('CourseID') or '').strip()
                            if cid and not cid.startswith(user_major):
                                continue
                        day = row.get('Day', '').strip()
                        slot = row.get('TimeSlot', '').strip()
                        if day and slot:
                            filled_slots_set.add(f"{day}-{slot}")
                stats['filled_slots'] = len(filled_slots_set)

                # Nếu vì lọc theo major mà về 0 (chưa có bản theo major), fallback: đếm toàn bộ
                if stats['total_classes'] == 0 and user_major in ('EE','ET'):
                    with open(schedule_path, 'r', encoding='utf-8-sig') as f2:
                        reader2 = csv.DictReader(f2)
                        for row in reader2:
                            stats['total_classes'] += 1
                            day = (row.get('Day') or '').strip()
                            slot = (row.get('TimeSlot') or '').strip()
                            if day in stats['classes_by_day']:
                                stats['classes_by_day'][day] += 1
                                if slot in ('1','2','3','4'):
                                    stats['classes_by_day_slot'][day][slot] += 1
                            room = (row.get('RoomAssigned') or '').strip()
                            if room:
                                rooms_set.add(room)
                            course = (row.get('CourseID') or '').strip()
                            if course:
                                courses_set.add(course)
                    stats['total_rooms'] = len(rooms_set)
                    stats['total_courses'] = len(courses_set)
                    filled_slots_set = set()
                    with open(schedule_path, 'r', encoding='utf-8-sig') as f3:
                        reader3 = csv.DictReader(f3)
                        for row in reader3:
                            day = row.get('Day', '').strip()
                            slot = row.get('TimeSlot', '').strip()
                            if day and slot:
                                filled_slots_set.add(f"{day}-{slot}")
                    stats['filled_slots'] = len(filled_slots_set)
            
            # Top recommendations từ ai_ranked_classes.csv
            ai_rank_path = resolve_major_path('ai_ranked_classes.csv')
            if os.path.exists(ai_rank_path):
                with open(ai_rank_path, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    top_recs = []
                    user_major = session.get('user_major')
                    for row in reader:
                        if user_major in ('EE','ET'):
                            cid0 = (row.get('CourseID') or row.get('Mã_HP') or '').strip()
                            if cid0 and not cid0.startswith(user_major):
                                continue
                        course_id = row.get('CourseID') or row.get('Mã_HP', '')
                        subject = row.get('SubjectName') or row.get('Tên_HP', '')
                        score = row.get('ai_score', '')
                        if course_id and subject:
                            top_recs.append({
                                'course': course_id,
                                'subject': subject[:50],  # Giới hạn độ dài
                                'score': float(score) if score else 0
                            })
                        if len(top_recs) >= 5:
                            break
                    stats['top_recommendations'] = top_recs

                    # Fallback khi sau lọc không còn gợi ý nào
                    if not stats['top_recommendations'] and user_major in ('EE','ET'):
                        f.seek(0)
                        reader_all = csv.DictReader(f)
                        top_recs = []
                        for row in reader_all:
                            course_id = row.get('CourseID') or row.get('Mã_HP', '')
                            subject = row.get('SubjectName') or row.get('Tên_HP', '')
                            score = row.get('ai_score', '')
                            if course_id and subject:
                                top_recs.append({
                                    'course': course_id,
                                    'subject': subject[:50],
                                    'score': float(score) if score else 0
                                })
                            if len(top_recs) >= 5:
                                break
                        stats['top_recommendations'] = top_recs
        except Exception as e:
            # Trả về stats mặc định nếu có lỗi
            pass
        
        return jsonify(stats)

    @app.route('/api/personal_stats')
    def api_personal_stats():
        """Thống kê cho cá nhân: hôm nay và cả tuần. Ưu tiên schedule_recommended.csv, rồi TKB_ca_nhan.csv"""
        import csv
        import datetime as dt

        def first_existing(paths):
            for p in paths:
                abs_p = project_path(p)
                if os.path.exists(abs_p):
                    return abs_p
            return None

        source = first_existing(['schedule_recommended.csv', 'TKB_ca_nhan.csv'])
        if not source:
            return jsonify({'has_data': False, 'today': [], 'counts_by_day': {}})

        # Map cột linh hoạt (VN/EN)
        with open(source, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            rows = list(reader)
        if not rows:
            return jsonify({'has_data': False, 'today': [], 'counts_by_day': {}})
        headers = rows[0]
        def col_idx(*names):
            for n in names:
                if n in headers:
                    return headers.index(n)
            return -1
        idx_day = col_idx('Day', 'Thứ')
        idx_time = col_idx('TimeSlot', 'Thời_gian')
        idx_subject = col_idx('SubjectName', 'Tên_HP')
        idx_course = col_idx('CourseID', 'Mã_HP')
        idx_room = col_idx('Room', 'Phòng', 'RoomAssigned')

        day_map = {'Mon':'Thứ 2', 'Tue':'Thứ 3', 'Wed':'Thứ 4', 'Thu':'Thứ 5', 'Fri':'Thứ 6', 'Sat':'Thứ 7'}
        reverse_day = {v:k for k,v in day_map.items()}

        # Hôm nay (theo hệ thống)
        today_py = dt.date.today().weekday()  # Mon=0
        today_key = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'][today_py]

        def normalize_time(s: str) -> str:
            s = (s or '').strip()
            if not s:
                return s
            if ':' in s:
                return s
            if '-' in s:
                parts = [p.strip() for p in s.split('-')]
                out = []
                for p in parts:
                    if len(p) == 4 and p.isdigit():
                        out.append(p[:2] + ':' + p[2:])
                    else:
                        out.append(p)
                return '-'.join(out)
            if len(s) == 4 and s.isdigit():
                return s[:2] + ':' + s[2:]
            return s

        counts = {'Mon':0,'Tue':0,'Wed':0,'Thu':0,'Fri':0,'Sat':0}
        today_list = []
        for r in rows[1:]:
            if max(idx_day, idx_time) < 0:
                continue
            day = (r[idx_day] if idx_day >= 0 else '').strip()
            # Chuẩn hoá day về Mon..Sat
            if day.startswith('Thứ'):
                day = reverse_day.get(day, day)
            time_slot = normalize_time(r[idx_time] if idx_time >= 0 else '')
            if not day or not time_slot:
                continue
            counts[day] = counts.get(day, 0) + 1
            if day == today_key:
                subject = (r[idx_subject] if idx_subject >= 0 else '') or (r[idx_course] if idx_course >= 0 else '')
                room = (r[idx_room] if idx_room >= 0 else '')
                today_list.append({'time': time_slot, 'subject': subject, 'room': room})

        # Sắp xếp lịch hôm nay theo giờ bắt đầu
        def start_minutes(t: str) -> int:
            a = (t.split('-')[0] if '-' in t else t)
            a = a.replace(':','')
            return int(a) if a.isdigit() else 0
        today_list.sort(key=lambda x: start_minutes(x['time']))

        return jsonify({'has_data': True, 'source': os.path.basename(source), 'today': today_list, 'counts_by_day': counts})

    @app.route('/preview')
    def preview():
        filename = request.args.get('file')
        page = int(request.args.get('page', 1))
        page_size = int(request.args.get('page_size', 50))
        if not filename or filename not in app.allowed_files:
            abort(400, description='file không hợp lệ')
        # Ưu tiên bản theo major (nếu có)
        path = resolve_major_path(filename)
        if not os.path.exists(path):
            abort(404)
        rows = []
        try:
            import csv
            with open(path, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                for row in reader:
                    rows.append(row)
        except Exception:
            # Nếu là JSON
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                # Hiển thị dạng bảng key/value cho JSON object - format dễ đọc
                def format_value(v):
                    """Chuyển giá trị JSON thành chuỗi dễ đọc"""
                    if v is None:
                        return '(trống)'
                    if isinstance(v, bool):
                        return 'Có' if v else 'Không'
                    if isinstance(v, (int, float)):
                        return str(v)
                    if isinstance(v, str):
                        return v
                    if isinstance(v, (list, tuple)):
                        if len(v) == 0:
                            return '(danh sách trống)'
                        if len(v) <= 3:
                            return ', '.join(format_value(x) for x in v)
                        return ', '.join(format_value(x) for x in v[:3]) + f' ... (tổng {len(v)} mục)'
                    if isinstance(v, dict):
                        if len(v) == 0:
                            return '(đối tượng trống)'
                        items = list(v.items())[:3]
                        parts = [f'{k}: {format_value(val)}' for k, val in items]
                        if len(v) > 3:
                            parts.append(f'... (tổng {len(v)} thuộc tính)')
                        return '; '.join(parts)
                    return str(v)
                
                if isinstance(data, dict):
                    rows = [['Thuộc tính', 'Giá trị']]
                    for k, v in data.items():
                        rows.append([str(k), format_value(v)])
                elif isinstance(data, list):
                    if len(data) == 0:
                        rows = [['Thông báo'], ['Danh sách trống']]
                    elif isinstance(data[0], dict):
                        # Nếu là danh sách object, hiển thị dạng bảng
                        all_keys = set()
                        for item in data:
                            if isinstance(item, dict):
                                all_keys.update(item.keys())
                        headers = ['STT'] + sorted(all_keys)
                        rows = [headers]
                        for idx, item in enumerate(data, 1):
                            row = [str(idx)]
                            for key in sorted(all_keys):
                                row.append(format_value(item.get(key, '')))
                            rows.append(row)
                    else:
                        rows = [['STT', 'Giá trị']]
                        for idx, item in enumerate(data, 1):
                            rows.append([str(idx), format_value(item)])
                else:
                    rows = [['Giá trị'], [format_value(data)]]
            except Exception as exc:
                abort(400, description=f'Không đọc được file: {exc}')

        total = len(rows)
        start = max(0, (page - 1) * page_size)
        end = min(total, start + page_size)
        headers = rows[0] if rows else []
        body = rows[1:] if rows else []
        page_rows = body[start:end]
        return render_template(
            'preview.html',
            filename=filename,
            headers=headers,
            rows=page_rows,
            page=page,
            page_size=page_size,
            total_rows=total - 1 if total > 0 else 0,
        )

    @app.route('/download')
    def download():
        filename = request.args.get('file')
        if not filename or filename not in app.allowed_files:
            abort(400, description='file không hợp lệ')
        path = resolve_major_path(filename)
        if not os.path.exists(path):
            abort(404)
        return send_file(path, as_attachment=True, download_name=filename)

    @app.route('/upload', methods=['GET', 'POST'])
    def upload():
        if request.method == 'GET':
            return render_template('upload.html', allowed=sorted(app.allowed_files))
        # POST
        target_name = request.form.get('target_name')
        file = request.files.get('file')
        if not target_name or target_name not in app.allowed_files:
            abort(400, description='target_name không hợp lệ')
        if not file:
            abort(400, description='Thiếu file upload')
        dest = project_path(target_name)
        file.save(dest)
        return jsonify({'ok': True, 'saved_to': target_name})

    @app.route('/constraints', methods=['GET', 'POST'])
    def constraints():
        path = project_path('constraints.json')
        if request.method == 'GET':
            if not os.path.exists(path):
                return render_template('constraints.html', constraints_text='{}')
            with open(path, 'r', encoding='utf-8') as f:
                return render_template('constraints.html', constraints_text=f.read())
        # POST
        text = request.form.get('constraints') or ''
        try:
            parsed = json.loads(text)
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(parsed, f, ensure_ascii=False, indent=2)
        except Exception as exc:
            abort(400, description=f'JSON không hợp lệ: {exc}')
        return jsonify({'ok': True})

    # Timetable pages
    @app.route('/schedule/new')
    def schedule_new():
        return render_template('schedule_new.html')

    @app.route('/timetable/school')
    def timetable_school():
        return render_template('timetable_school.html')

    @app.route('/timetable/student')
    def timetable_student():
        return render_template('timetable_student.html')

    @app.route('/timetable/teacher')
    def timetable_teacher():
        return render_template('timetable_teacher.html')

    @app.route('/home/teacher')
    def home_teacher():
        return render_template('home_teacher.html')

    @app.route('/home/student')
    def home_student():
        return render_template('home_student.html')

    return app


if __name__ == '__main__':
    app = create_app()
    # Chạy chỉ trên localhost
    threading.Timer(1.0, lambda: webbrowser.open('http://127.0.0.1:5000/')).start()
    app.run(host='127.0.0.1', port=5000, debug=True)


