#!/usr/bin/env python3
"""Fix config.ps1 with proper UTF-8 BOM encoding.

Usage:
    python fix_config.py [target_config_ps1_path]

If no path is given, defaults to <this_script_dir>/config.ps1.
"""
import sys
from pathlib import Path

# 动态定位：基于脚本自身目录，不硬编码任何绝对路径
SCRIPT_DIR = Path(__file__).resolve().parent

config_content = """\
# ============================================================
# 智联招聘简历批量下载 — 外部配置文件（已弃用，请改用 config.json）
# ============================================================
# 此文件不再被 run.ps1 使用。run.ps1 直接读取 config.json。
# 如需使用，请通过命令行参数传入或编辑 config.json。
# ============================================================

$Config = @{
    Url            = ""
    JobName        = ""
    DownloadDir    = ""
    DownloadCount  = 5
    FileFormat     = "word"
    WebBridgeUrl = "http://127.0.0.1:10086/command"
    Session      = "resume-screening"
    SaveLocalX   = 1079
    SaveLocalY   = 142
    SaveConfirmX = 996
    SaveConfirmY = 561
    MaskCloseX   = 30
    MaskCloseY   = 300
    CloseWaitMs        = 700
    ClickWaitMs        = 800
    MouseMoveWaitMs    = 900
    PressReleaseWaitMs = 400
    DialogCheckWaitMs  = 2000
    DownloadWaitMs     = 6000
    ScrollWaitMs       = 1200
    DownloadFilter = "*智联简历*"
    DownloadSource = "$env:USERPROFILE\\Downloads"
}
"""

if __name__ == "__main__":
    if len(sys.argv) > 1:
        target = Path(sys.argv[1])
    else:
        target = SCRIPT_DIR / "config.ps1"

    with open(target, "wb") as f:
        f.write(b"\xef\xbb\xbf")  # UTF-8 BOM
        f.write(config_content.encode("utf-8"))

    print(f"OK: {target} written with UTF-8 BOM")
