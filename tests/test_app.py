import io
import re
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Inches
from werkzeug.security import generate_password_hash

from app import create_app


def deck(label, wide=True):
    p = Presentation()
    if wide:
        p.slide_width, p.slide_height = Inches(13.333), Inches(7.5)
    s = p.slides.add_slide(p.slide_layouts[5])
    s.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1)).text = label
    b = io.BytesIO()
    p.save(b)
    return b.getvalue()


@pytest.fixture
def client(tmp_path):
    app = create_app({"TESTING": True, "DATA_DIR": str(tmp_path), "SECRET_KEY": "synthetic-session-key", "ADMIN_USERNAME": "admin", "ADMIN_PASSWORD_HASH": generate_password_hash("synthetic-password")})
    return app.test_client()


def token(html):
    return re.search(rb'name="csrf_token" value="([^"]+)"', html).group(1).decode()


def login(client):
    t = token(client.get('/admin/login').data)
    return client.post('/admin/login', data={"username": "admin", "password": "synthetic-password", "csrf_token": t}, follow_redirects=True)


def upload(client, name="张三", date="20261102", content=None, overwrite=None):
    t = token(client.get('/').data)
    fields = {"name": name, "date": date, "file": (io.BytesIO(content or deck(name)), "talk.pptx"), "csrf_token": t}
    if overwrite:
        fields["overwrite_token"] = overwrite
    return client.post('/upload', data=fields, content_type="multipart/form-data")


def test_upload_validation_and_overwrite(client):
    assert upload(client).status_code == 200
    conflict = upload(client, content=deck("new"))
    assert conflict.status_code == 409
    m = re.search(r'name="overwrite_token" value="([^"]+)"', conflict.get_data(as_text=True))
    assert m
    assert upload(client, content=deck("new"), overwrite=m.group(1)).status_code == 200
    assert upload(client, date="20260230").status_code == 400
    assert upload(client, name="../evil").status_code == 400
    assert upload(client, content=b"not a pptx").status_code == 400


def test_conflict_does_not_change_existing_and_persists(tmp_path):
    config = {"TESTING": True, "DATA_DIR": str(tmp_path), "SECRET_KEY": "synthetic-session-key", "ADMIN_USERNAME": "admin", "ADMIN_PASSWORD_HASH": generate_password_hash("synthetic-password")}
    c = create_app(config).test_client()
    assert upload(c, content=deck('original')).status_code == 200
    assert upload(c, content=deck('replacement')).status_code == 409
    c = create_app(config).test_client()
    login(c)
    t = token(c.get('/admin?date=20261102').data)
    merged = c.post('/admin/merge', data={'date': '20261102', 'csrf_token': t})
    assert merged.status_code == 200
    slide = Presentation(io.BytesIO(merged.data)).slides[0]
    assert any(shape.has_text_frame and shape.text == 'original' for shape in slide.shapes)


def test_history_skips_absent_and_appends_new(client):
    upload(client, '张三')
    upload(client, '李四')
    login(client)
    t = token(client.get('/admin?date=20261102').data)
    client.post('/admin/history/save', data={'date': '20261102', 'label': '常用', 'csrf_token': t})
    upload(client, '李四', date='20261109')
    upload(client, '王五', date='20261109')
    assert client.post('/admin/history/apply', data={'date': '20261109', 'label': '常用', 'csrf_token': t}).status_code == 302
    page = client.get('/admin?date=20261109').get_data(as_text=True)
    assert page.index('<strong>李四</strong>') < page.index('<strong>王五</strong>')
    assert '缺席成员已跳过' in page


def test_admin_auth_csrf_download_delete_and_order(client):
    upload(client, "张三")
    upload(client, "李四")
    assert client.get('/admin').status_code == 302
    assert client.get('/admin/zip?date=20261102').status_code == 302
    assert login(client).status_code == 200
    page = client.get('/admin?date=20261102').get_data(as_text=True)
    assert "张三" in page and "李四" in page
    assert client.post('/admin/order', data={"date": "20261102", "name": "李四", "direction": "up"}).status_code == 400
    t = token(page.encode())
    assert client.post('/admin/order', data={"date": "20261102", "name": "李四", "direction": "up", "csrf_token": t}).status_code == 302
    assert client.get('/admin/zip?date=20261102').status_code == 200
    assert client.post('/admin/delete', data={"date": "20261102", "names": "张三", "csrf_token": t}).status_code == 302
    assert "张三" not in client.get('/admin?date=20261102').get_data(as_text=True)
    assert client.post('/admin/delete', data={"date": "20261102", "all": "1", "csrf_token": t}).status_code == 302
    assert "李四" not in client.get('/admin?date=20261102').get_data(as_text=True)
    assert client.get('/admin/zip?date=20261102').status_code == 404


def test_upload_requires_csrf(client):
    response = client.post('/upload', data={"name": "张三", "date": "20261102", "file": (io.BytesIO(deck('x')), 'x.pptx')}, content_type='multipart/form-data')
    assert response.status_code == 400


def test_merge_and_history(client):
    upload(client, "张三")
    upload(client, "李四")
    login(client)
    t = token(client.get('/admin?date=20261102').data)
    assert client.post('/admin/history/save', data={"date": "20261102", "label": "常用", "csrf_token": t}).status_code == 302
    r = client.post('/admin/merge', data={"date": "20261102", "csrf_token": t})
    assert r.status_code == 200
    p = Presentation(io.BytesIO(r.data))
    assert len(p.slides) == 2
    assert client.post('/admin/history/delete', data={"label": "常用", "csrf_token": t}).status_code == 302


def test_merge_form_explicitly_forwards_notes_normalization(client, monkeypatch):
    upload(client, '张三')
    login(client)
    page = client.get('/admin?date=20261102').get_data(as_text=True)
    assert 'name="normalize_notes_master"' in page
    assert '统一备注母版（保留备注内容，备注排版可能变化）' in page
    assert not re.search(r'<input[^>]+name="normalize_notes_master"[^>]+checked', page)
    import app as app_module
    actual = app_module.merge_pptx
    seen = []
    def recording(files, normalize_notes_master=False):
        seen.append(normalize_notes_master)
        return actual(files, normalize_notes_master=normalize_notes_master)
    monkeypatch.setattr(app_module, 'merge_pptx', recording)
    t = token(page.encode())
    assert client.post('/admin/merge', data={'date': '20261102', 'csrf_token': t}).status_code == 200
    assert client.post('/admin/merge', data={'date': '20261102', 'csrf_token': t, 'normalize_notes_master': '1'}).status_code == 200
    assert seen == [False, True]


def test_login_throttle(client):
    t = token(client.get('/admin/login').data)
    for _ in range(5):
        client.post('/admin/login', data={"username": "admin", "password": "wrong", "csrf_token": t})
    assert client.post('/admin/login', data={"username": "admin", "password": "synthetic-password", "csrf_token": t}).status_code == 429


def test_compact_upload_page_has_accessible_file_picker(client):
    page = client.get('/').get_data(as_text=True)
    assert '把这一期的研究' not in page
    assert 'GROUP MEETING ARCHIVE' not in page
    assert 'class="hero"' not in page
    assert '提交组会 PPT' in page
    assert re.search(r'<input[^>]+id="upload-file"[^>]+type="file"[^>]+accept="\.pptx,\.ppt"', page)
    assert '<label for="upload-file"' in page
    assert '选择 PPT 文件' in page
