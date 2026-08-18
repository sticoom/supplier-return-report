"""T8：部署模块 —— 参照服务器 /opt/fund-dashboard 既有 systemd 模式生成部署产物。"""
from __future__ import annotations

from pathlib import Path

HOST = "120.25.100.51"
PORT = 8300
REMOTE_DIR = "/opt/supplier-return-report"
SERVICE = "supplier-return"
REPO_ROOT = Path(__file__).resolve().parent.parent


def parse_sections(unit: str) -> dict[str, list[str]]:
    """把 unit 文本解析成 {节名: [该节行]}，便于断言。"""
    sections: dict[str, list[str]] = {}
    current = ""
    for line in unit.splitlines():
        line = line.strip()
        if line.startswith("[") and line.endswith("]"):
            current = line
            sections[current] = []
        elif line and current:
            sections[current].append(line)
    return sections


def test_systemd_unit_follows_fund_dashboard_pattern():
    """unit 内容：systemd 管理 /opt/supplier-return-report 下的 uvicorn 8300。"""
    from deploy.service import systemd_unit

    unit = systemd_unit()
    sections = parse_sections(unit)

    assert "[Unit]" in sections and "[Service]" in sections and "[Install]" in sections
    assert any(l.startswith("Description=") for l in sections["[Unit]"])
    assert any(l.startswith("After=") and "network.target" in l for l in sections["[Unit]"])

    svc = sections["[Service]"]
    assert f"WorkingDirectory={REMOTE_DIR}" in svc
    exec_start = next(l for l in svc if l.startswith("ExecStart="))
    # venv 里的 uvicorn 启动 FastAPI 应用，监听 0.0.0.0:8300
    assert exec_start == (
        f"ExecStart={REMOTE_DIR}/.venv/bin/uvicorn web.main:app "
        f"--host 0.0.0.0 --port {PORT}"
    )
    # fund-dashboard 模式：服务挂了自动拉起（而不是 nohup 一次性）
    assert "Restart=always" in svc

    assert f"WantedBy=multi-user.target" in sections["[Install]"]


def test_install_and_redeploy_commands():
    """服务器侧安装/更新命令：venv+pip+systemd 闭环，禁止 nohup/pkill 反模式。"""
    from deploy.service import install_commands, redeploy_commands

    install = install_commands()
    joined = "\n".join(install)
    # 在远端目录建 venv 并装生产依赖
    assert f"cd {REMOTE_DIR}" in joined and "python3 -m venv .venv" in joined
    assert ".venv/bin/pip install -r requirements.txt" in joined
    # unit 落盘 + systemd 闭环
    assert "daemon-reload" in joined and f"enable --now {SERVICE}" in joined
    # fund-dashboard 模式：systemd 接管生命周期，不用 nohup/pkill/screen
    for bad in ("nohup", "pkill", "screen "):
        assert bad not in joined

    redeploy = "\n".join(redeploy_commands())
    # 代码更新后：重启即生效
    assert f"systemctl restart {SERVICE}" in redeploy
    assert "daemon-reload" not in redeploy  # unit 未变时无需 reload


def test_sync_commands_copy_code_not_data():
    """本机→服务器同步清单：代码+依赖清单进 /opt，data/ 与虚拟环境不上传。"""
    from deploy.service import sync_commands

    joined = "\n".join(sync_commands())
    assert "scp" in joined and f"root@{HOST}:{REMOTE_DIR}" in joined
    for path in ("engine", "web", "requirements.txt", "pyproject.toml", "deploy"):
        assert path in joined
    assert "data" not in joined.replace(REMOTE_DIR, "")  # data/ 不上传
    assert ".venv" not in joined
    assert "tests" not in joined  # 服务器只跑生产，不需要测试目录


def test_wait_healthy_against_real_uvicorn(tmp_path):
    """验收工具：真实 uvicorn 起本应用 → wait_healthy 探活成功；死端口探活失败。"""
    import socket
    import threading
    import time

    import uvicorn
    from deploy.service import wait_healthy
    from web.main import create_app

    app = create_app(tmp_path)  # 数据目录指向 tmp，不碰真实 data/
    config = uvicorn.Config(app, host="127.0.0.1", port=0, log_level="warning")
    server = uvicorn.Server(config)
    thread = threading.Thread(target=server.run, daemon=True)
    thread.start()
    for _ in range(200):  # 等 uvicorn 完成启动
        if server.started:
            break
        time.sleep(0.05)
    assert server.started
    port = server.servers[0].sockets[0].getsockname()[1]

    # 活着的服务：GET / 200 → True
    assert wait_healthy(f"http://127.0.0.1:{port}/", timeout=10) is True

    # 已关闭的端口：超时后 False（不抛异常）
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    dead_port = s.getsockname()[1]
    s.close()
    assert wait_healthy(f"http://127.0.0.1:{dead_port}/", timeout=1) is False

    server.should_exit = True
    thread.join(timeout=5)


def test_cli_prints_unit_and_plan():
    """CLI：`python -m deploy.service unit` 输出与 systemd_unit() 一致；plan 输出步骤清单。"""
    import os
    import subprocess
    import sys

    from deploy.service import install_commands, sync_commands, systemd_unit

    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}  # 子进程输出统一 UTF-8
    r = subprocess.run([sys.executable, "-m", "deploy.service", "unit"],
                       capture_output=True, text=True, encoding="utf-8",
                       cwd=REPO_ROOT, env=env)
    assert r.returncode == 0, r.stderr
    assert r.stdout == systemd_unit()

    r = subprocess.run([sys.executable, "-m", "deploy.service", "plan"],
                       capture_output=True, text=True, encoding="utf-8",
                       cwd=REPO_ROOT, env=env)
    assert r.returncode == 0, r.stderr
    for cmd in sync_commands() + install_commands():
        assert cmd in r.stdout
