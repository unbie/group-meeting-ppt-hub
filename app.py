from __future__ import annotations

import io
import json
import os
import re
import secrets
import shutil
import sqlite3
import subprocess
import tempfile
import time
import unicodedata
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

from flask import Flask, abort, flash, redirect, render_template, request, send_file, session, url_for
from itsdangerous import BadSignature, URLSafeSerializer
from pptx import Presentation
from werkzeug.security import check_password_hash

from merge import MergeError, merge_pptx

MAX_UPLOAD = 50 * 1024 * 1024
MAX_EXPANDED = 300 * 1024 * 1024
MAX_PARTS = 3000


def valid_date(value):
    if not re.fullmatch(r'\d{8}', value or ''):
        raise ValueError('日期须为 YYYYMMDD 八位数字')
    try:
        datetime.strptime(value, '%Y%m%d')
    except ValueError:
        raise ValueError('请输入真实的公历日期') from None
    return value


def valid_name(value):
    value = unicodedata.normalize('NFC', (value or '').strip())
    if not 1 <= len(value) <= 40 or any(unicodedata.category(c).startswith('C') or c in '/\\<>:"|?*' for c in value) or value in {'.', '..'}:
        raise ValueError('姓名须为 1–40 个字符，不能含路径或控制字符')
    return value


def validate_pptx(raw, max_bytes=MAX_UPLOAD):
    if len(raw) > max_bytes or len(raw) < 50:
        raise ValueError('文件大小不符合要求')
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            infos = z.infolist()
            if len(infos) > MAX_PARTS or sum(i.file_size for i in infos) > MAX_EXPANDED:
                raise ValueError('PPTX 解压后过大或部件过多')
            names = [i.filename for i in infos]
            if len(names) != len(set(names)) or any(n.startswith('/') or '\\' in n or '..' in n.split('/') for n in names):
                raise ValueError('PPTX 含不安全的文件路径')
            if any(i.flag_bits & 1 or (i.compress_size and i.file_size / i.compress_size > 1000) for i in infos):
                raise ValueError('PPTX 包含加密或异常压缩部件')
            if not {'[Content_Types].xml', 'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels'} <= set(names):
                raise ValueError('文件不是完整的 PPTX')
            if any('vbaproject' in n.lower() or 'activex/' in n.lower() for n in names):
                raise ValueError('不接受宏或 ActiveX 部件')
            for name in names:
                if name.endswith(('.xml', '.rels')):
                    data = z.read(name)
                    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
                        raise ValueError('Office XML 含不安全声明')
                    ET.fromstring(data)
        p = Presentation(io.BytesIO(raw))
        if len(p.slides) == 0:
            raise ValueError('演示文稿没有幻灯片')
        return len(p.slides)
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        raise ValueError('文件不是有效的 PPTX') from exc


def create_app(overrides=None):
    app = Flask(__name__)
    local_config = Path(__file__).parent / 'config.json'
    saved = json.loads(local_config.read_text(encoding='utf-8')) if local_config.exists() else {}
    app.config.update(DATA_DIR=os.environ.get('PPT_HUB_DATA_DIR', str(Path(__file__).parent / 'data')),
                      SECRET_KEY=os.environ.get('PPT_HUB_SECRET_KEY') or saved.get('secret_key'),
                      ADMIN_USERNAME=os.environ.get('PPT_HUB_ADMIN_USER', 'admin'),
                      ADMIN_PASSWORD_HASH=os.environ.get('PPT_HUB_ADMIN_PASSWORD_HASH') or saved.get('admin_password_hash'),
                      SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE='Lax',
                      SESSION_COOKIE_SECURE=os.environ.get('PPT_HUB_HTTPS') == '1',
                      MAX_CONTENT_LENGTH=MAX_UPLOAD + 1024 * 1024)
    if overrides:
        app.config.update(overrides)
    if not app.config['SECRET_KEY'] or not app.config['ADMIN_PASSWORD_HASH']:
        raise RuntimeError('请先运行 setup.py 设置管理员密码和会话密钥')
    if not app.config['TESTING'] and (len(app.config['SECRET_KEY']) < 32 or not app.config['ADMIN_PASSWORD_HASH'].startswith(('scrypt:', 'pbkdf2:'))):
        raise RuntimeError('密钥或密码哈希无效；请使用 setup.py 生成配置')
    data_dir = Path(app.config['DATA_DIR']).resolve()
    files_dir = data_dir / 'uploads'
    files_dir.mkdir(parents=True, exist_ok=True)
    db_path = data_dir / 'hub.sqlite3'

    def db():
        conn = sqlite3.connect(db_path, timeout=30)
        conn.row_factory = sqlite3.Row
        conn.execute('PRAGMA foreign_keys=ON')
        return conn

    with db() as c:
        c.executescript('''CREATE TABLE IF NOT EXISTS submissions (date TEXT NOT NULL, name TEXT NOT NULL, file_id TEXT NOT NULL, original_name TEXT NOT NULL, slides INTEGER NOT NULL, updated_at TEXT NOT NULL, PRIMARY KEY(date,name));
        CREATE TABLE IF NOT EXISTS ordering (date TEXT NOT NULL, name TEXT NOT NULL, position INTEGER NOT NULL, PRIMARY KEY(date,name));
        CREATE TABLE IF NOT EXISTS histories (label TEXT PRIMARY KEY, names TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS login_attempts (ip TEXT PRIMARY KEY, failures INTEGER NOT NULL, first_at INTEGER NOT NULL);''')
        known = {row[0] for row in c.execute('SELECT file_id FROM submissions')}
    for path in files_dir.iterdir():
        if path.name.endswith('.deleting'):
            original = files_dir / path.name[:-9]
            if original.name in known and not original.exists():
                os.replace(path, original)
            else:
                path.unlink(missing_ok=True)
        elif path.name.endswith('.tmp') or (path.name.endswith('.pptx') and path.name not in known):
            path.unlink(missing_ok=True)

    signer = URLSafeSerializer(app.config['SECRET_KEY'], salt='overwrite')

    def csrf():
        if 'csrf' not in session:
            session['csrf'] = secrets.token_urlsafe(32)
        return session['csrf']

    app.jinja_env.globals['csrf_token'] = csrf

    @app.before_request
    def protect():
        if request.method == 'POST' and not secrets.compare_digest(request.form.get('csrf_token', ''), csrf()):
            abort(400, '页面已过期，请刷新后重试（CSRF）')
        if request.path.startswith('/admin') and request.path not in {'/admin/login'} and not session.get('admin'):
            return redirect(url_for('login'))

    def rows_for(date):
        with db() as c:
            rows = c.execute('''SELECT s.*, COALESCE(o.position,1000000) AS position FROM submissions s LEFT JOIN ordering o ON s.date=o.date AND s.name=o.name WHERE s.date=? ORDER BY position, s.updated_at, s.name''', (date,)).fetchall()
        return rows

    def converted(raw, filename):
        extension = Path(filename).suffix.lower()
        if extension == '.pptx':
            return raw
        if extension != '.ppt':
            raise ValueError('仅接受 .pptx 或 .ppt 文件')
        soffice = shutil.which('soffice') or shutil.which('libreoffice')
        if not soffice:
            raise ValueError('旧版 .ppt 需要 LibreOffice 转换；当前未安装，请先另存为 .pptx')
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'input.ppt'
            source.write_bytes(raw)
            result = subprocess.run([soffice, '-env:UserInstallation=file:///' + folder.replace('\\', '/') + '/lo-profile', '--headless', '--convert-to', 'pptx', '--outdir', folder, str(source)], capture_output=True, timeout=90)
            output = Path(folder) / 'input.pptx'
            if result.returncode or not output.exists():
                raise ValueError('LibreOffice 无法转换该 .ppt 文件')
            return output.read_bytes()

    @app.get('/')
    def home():
        return render_template('upload.html')

    @app.post('/upload')
    def upload():
        try:
            name, date = valid_name(request.form.get('name')), valid_date(request.form.get('date'))
            f = request.files.get('file')
            if not f or not f.filename:
                raise ValueError('请选择 PPT 文件')
            filename = Path(f.filename).name
            if Path(filename).suffix.lower() not in {'.ppt', '.pptx'}:
                raise ValueError('仅接受 .pptx 或 .ppt 文件')
            raw = f.read(MAX_UPLOAD + 1)
            if len(raw) > MAX_UPLOAD:
                raise ValueError('文件超过 50 MB 限制')
            raw = converted(raw, filename)
            slides = validate_pptx(raw)
            conn = db()
            new_id = uuid.uuid4().hex + '.pptx'
            destination = files_dir / new_id
            try:
                conn.execute('BEGIN IMMEDIATE')
                old = conn.execute('SELECT file_id FROM submissions WHERE date=? AND name=?', (date, name)).fetchone()
                if old:
                    token = request.form.get('overwrite_token', '')
                    try:
                        approved = signer.loads(token)
                    except BadSignature:
                        approved = None
                    if approved != [date, name, old['file_id']]:
                        conn.rollback()
                        confirmation = signer.dumps([date, name, old['file_id']])
                        return render_template('conflict.html', name=name, date=date, overwrite_token=confirmation), 409
                temp = files_dir / (new_id + '.tmp')
                try:
                    temp.write_bytes(raw)
                    os.replace(temp, destination)
                    conn.execute('INSERT OR REPLACE INTO submissions VALUES (?,?,?,?,?,?)', (date, name, new_id, filename, slides, datetime.now().isoformat(timespec='seconds')))
                    conn.execute('INSERT OR IGNORE INTO ordering VALUES (?,?,(SELECT COALESCE(MAX(position),-1)+1 FROM ordering WHERE date=?))', (date, name, date))
                    conn.commit()
                except Exception:
                    conn.rollback()
                    temp.unlink(missing_ok=True)
                    destination.unlink(missing_ok=True)
                    raise
                if old:
                    (files_dir / old['file_id']).unlink(missing_ok=True)
            finally:
                conn.close()
            return render_template('success.html', name=name, date=date)
        except (ValueError, OSError, subprocess.TimeoutExpired) as exc:
            return render_template('upload.html', error=str(exc)), 400

    @app.route('/admin/login', methods=['GET', 'POST'])
    def login():
        if request.method == 'GET':
            return render_template('login.html')
        ip = request.remote_addr or 'local'
        now = int(time.time())
        with db() as c:
            c.execute('BEGIN IMMEDIATE')
            row = c.execute('SELECT * FROM login_attempts WHERE ip=?', (ip,)).fetchone()
            failures = row['failures'] if row and now - row['first_at'] < 900 else 0
            first = row['first_at'] if failures else now
            if failures >= 5:
                return render_template('login.html', error='尝试过多，请 15 分钟后重试'), 429
            if request.form.get('username') == app.config['ADMIN_USERNAME'] and check_password_hash(app.config['ADMIN_PASSWORD_HASH'], request.form.get('password', '')):
                c.execute('DELETE FROM login_attempts WHERE ip=?', (ip,))
                session.clear(); session['admin'] = True
                return redirect(url_for('admin'))
            c.execute('INSERT OR REPLACE INTO login_attempts VALUES (?,?,?)', (ip, failures + 1, first))
        return render_template('login.html', error='账号或密码错误'), 401

    @app.post('/admin/logout')
    def logout():
        session.clear()
        return redirect(url_for('home'))

    @app.get('/admin')
    def admin():
        date = request.args.get('date', datetime.now().strftime('%Y%m%d'))
        try:
            valid_date(date)
        except ValueError as exc:
            return render_template('admin.html', date=date, rows=[], histories=[], error=str(exc)), 400
        with db() as c:
            histories = c.execute('SELECT label,names FROM histories ORDER BY label').fetchall()
        return render_template('admin.html', date=date, rows=rows_for(date), histories=histories)

    def date_form():
        return valid_date(request.form.get('date'))

    @app.get('/admin/zip')
    def download_zip():
        try:
            date = valid_date(request.args.get('date'))
        except ValueError:
            abort(400)
        with db() as c:
            c.execute('BEGIN IMMEDIATE')
            rows = c.execute('''SELECT s.*, COALESCE(o.position,1000000) AS position FROM submissions s LEFT JOIN ordering o ON s.date=o.date AND s.name=o.name WHERE s.date=? ORDER BY position, s.updated_at, s.name''', (date,)).fetchall()
            if not rows:
                abort(404)
            contents = [(r, (files_dir / r['file_id']).read_bytes()) for r in rows]
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as z:
            for row, content in contents:
                z.writestr(f"{date}_{row['name']}.pptx", content)
        buffer.seek(0)
        return send_file(buffer, as_attachment=True, download_name=f'{date}.zip', mimetype='application/zip')

    @app.post('/admin/delete')
    def delete():
        try:
            date = date_form()
            names = request.form.getlist('names')
            if request.form.get('all') == '1':
                names = [r['name'] for r in rows_for(date)]
            if not names:
                raise ValueError('请选择要删除的文件')
            conn = db(); moved = []
            try:
                conn.execute('BEGIN IMMEDIATE')
                for name in names:
                    row = conn.execute('SELECT file_id FROM submissions WHERE date=? AND name=?', (date, name)).fetchone()
                    if row:
                        path = files_dir / row['file_id']; trash = files_dir / (row['file_id'] + '.deleting')
                        os.replace(path, trash); moved.append((path, trash))
                        conn.execute('DELETE FROM submissions WHERE date=? AND name=?', (date, name))
                conn.commit()
            except Exception:
                conn.rollback()
                for path, trash in reversed(moved):
                    os.replace(trash, path)
                raise
            finally:
                conn.close()
            for path, trash in moved:
                trash.unlink(missing_ok=True)
            flash('已删除所选文件')
        except ValueError as exc:
            flash(str(exc))
        return redirect(url_for('admin', date=request.form.get('date')))

    @app.post('/admin/order')
    def order():
        date = date_form(); names = [r['name'] for r in rows_for(date)]
        name = request.form.get('name'); direction = request.form.get('direction')
        if name not in names or direction not in {'up', 'down'}:
            abort(400)
        i = names.index(name); j = i + (-1 if direction == 'up' else 1)
        if 0 <= j < len(names):
            names[i], names[j] = names[j], names[i]
        with db() as c:
            for i, n in enumerate(names):
                c.execute('INSERT OR REPLACE INTO ordering VALUES (?,?,?)', (date, n, i))
        return redirect(url_for('admin', date=date))

    @app.post('/admin/history/save')
    def history_save():
        date = date_form(); label = (request.form.get('label') or '').strip()
        if not 1 <= len(label) <= 40:
            abort(400)
        names = [r['name'] for r in rows_for(date)]
        import json
        with db() as c:
            c.execute('INSERT OR REPLACE INTO histories VALUES (?,?)', (label, json.dumps(names, ensure_ascii=False)))
        return redirect(url_for('admin', date=date))

    @app.post('/admin/history/apply')
    def history_apply():
        import json
        date = date_form(); label = request.form.get('label')
        with db() as c:
            row = c.execute('SELECT names FROM histories WHERE label=?', (label,)).fetchone()
            if not row: abort(404)
            saved = json.loads(row['names'])
            present = [r['name'] for r in rows_for(date)]
            ordered = [n for n in saved if n in present] + [n for n in present if n not in saved]
            for i, n in enumerate(ordered):
                c.execute('INSERT OR REPLACE INTO ordering VALUES (?,?,?)', (date, n, i))
        flash('已应用历史顺序；缺席成员已跳过，新成员排在末尾')
        return redirect(url_for('admin', date=date))

    @app.post('/admin/history/delete')
    def history_delete():
        with db() as c:
            c.execute('DELETE FROM histories WHERE label=?', (request.form.get('label'),))
        return redirect(url_for('admin', date=request.form.get('date', datetime.now().strftime('%Y%m%d'))))

    @app.post('/admin/merge')
    def merge():
        try:
            date = date_form()
            with db() as c:
                c.execute('BEGIN IMMEDIATE')
                rows = c.execute('''SELECT s.*, COALESCE(o.position,1000000) AS position FROM submissions s LEFT JOIN ordering o ON s.date=o.date AND s.name=o.name WHERE s.date=? ORDER BY position, s.updated_at, s.name''', (date,)).fetchall()
                source = [(files_dir / r['file_id']).read_bytes() for r in rows]
            expected = sum(r['slides'] for r in rows)
            result = merge_pptx(source, normalize_notes_master=request.form.get('normalize_notes_master') == '1')
            if validate_pptx(result, max_bytes=500 * 1024 * 1024) != expected:
                raise MergeError('合并后页数不符，已取消下载')
            return send_file(io.BytesIO(result), as_attachment=True, download_name=f'{date}.pptx', mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')
        except (MergeError, ValueError, OSError) as exc:
            return render_template('admin.html', date=request.form.get('date'), rows=rows_for(request.form.get('date', '')), histories=[], error=str(exc)), 400

    @app.errorhandler(413)
    def too_large(_):
        return render_template('upload.html', error='文件超过 50 MB 限制'), 413

    return app


if __name__ == '__main__':
    from waitress import serve
    serve(create_app(), host=os.environ.get('PPT_HUB_HOST', '127.0.0.1'), port=int(os.environ.get('PPT_HUB_PORT', '1410')))
