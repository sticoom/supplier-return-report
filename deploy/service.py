"""T8 部署模块：生成 systemd unit 与部署/验收命令（参照 /opt/fund-dashboard 既有模式）。

用法：
    python -m deploy.service unit      # 打印 systemd unit 文件内容
    python -m deploy.service plan      # 打印部署步骤命令清单
"""
from __future__ import annotations

import time

import httpx

HOST = "120.25.100.51"
PORT = 8300
REMOTE_DIR = "/opt/supplier-return-report"
SERVICE = "supplier-return"
BASE_URL = f"http://{HOST}:{PORT}/"
UNIT_PATH = f"/etc/systemd/system/{SERVICE}.service"


def systemd_unit() -> str:
    """供应商退货统计服务的 systemd unit（与 fund-dashboard 同模式）。"""
    return f"""[Unit]
Description=供应商质量退货金额统计 (supplier-return-report)
After=network.target

[Service]
Type=simple
WorkingDirectory={REMOTE_DIR}
ExecStart={REMOTE_DIR}/.venv/bin/uvicorn web.main:app --host 0.0.0.0 --port {PORT}
Restart=always
RestartSec=3

[Install]
WantedBy=multi-user.target
"""


def sync_commands() -> list[str]:
    """本机 → 服务器同步代码（data/ 与 tests/ 不上传）。"""
    return [
        f'ssh root@{HOST} "mkdir -p {REMOTE_DIR}"',
        f"scp -r engine web deploy root@{HOST}:{REMOTE_DIR}/",
        f"scp requirements.txt pyproject.toml root@{HOST}:{REMOTE_DIR}/",
    ]


def install_commands() -> list[str]:
    """服务器侧首次安装：venv + 依赖 + unit 落盘 + systemd 接管。"""
    return [
        f"cd {REMOTE_DIR}",
        "python3 -m venv .venv",
        ".venv/bin/pip install -r requirements.txt",
        f".venv/bin/python -m deploy.service unit > {UNIT_PATH}",
        "systemctl daemon-reload",
        f"systemctl enable --now {SERVICE}",
    ]


def redeploy_commands() -> list[str]:
    """代码更新后：重装依赖（若有变化）并重启服务。"""
    return [
        f"cd {REMOTE_DIR} && .venv/bin/pip install -r requirements.txt",
        f"systemctl restart {SERVICE}",
    ]


def wait_healthy(url: str, timeout: float = 30.0, interval: float = 0.5) -> bool:
    """轮询 GET url 直到 200 或超时；用于部署后的线上网页验收。"""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            if httpx.get(url, timeout=2.0, trust_env=False).status_code == 200:
                return True   # trust_env=False：本机健康检查不走系统代理
        except httpx.HTTPError:
            pass
        time.sleep(interval)
    return False


def main(argv: list[str] | None = None) -> int:
    """CLI：unit 打印 unit 文件；plan 打印部署步骤（本机同步 + 服务器安装）。"""
    import sys

    arg = (argv or sys.argv[1:])[0] if (argv or sys.argv[1:]) else ""
    if arg == "unit":
        sys.stdout.write(systemd_unit())
        return 0
    if arg == "plan":
        for i, cmd in enumerate(sync_commands() + install_commands(), 1):
            print(f"{i}. {cmd}")
        return 0
    print(__doc__, file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
