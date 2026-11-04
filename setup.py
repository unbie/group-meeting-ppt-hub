"""Create an ignored local administrator configuration without printing secrets."""
import getpass
import json
import os
import secrets
import sys
from pathlib import Path

from werkzeug.security import generate_password_hash

docker = '--docker' in sys.argv
target = Path(__file__).parent / ('.env' if docker else 'config.json')
if target.exists():
    raise SystemExit(f'{target.name} 已存在；如需重设，请先备份并自行移走旧文件。')
password = getpass.getpass('设置管理员密码（至少 12 个字符）：')
confirm = getpass.getpass('再次输入管理员密码：')
if len(password) < 12 or password != confirm:
    raise SystemExit('密码长度不足或两次输入不一致。')
secret = secrets.token_urlsafe(48)
password_hash = generate_password_hash(password)
fd = os.open(target, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
with os.fdopen(fd, 'w', encoding='utf-8') as f:
    if docker:
        f.write('PPT_HUB_SECRET_KEY=' + secret + "\nPPT_HUB_ADMIN_PASSWORD_HASH='" + password_hash + "'\nPPT_HUB_ADMIN_USER=admin\n")
    else:
        json.dump({'secret_key': secret, 'admin_password_hash': password_hash}, f)
print(f'已创建本地 {target.name}。')
