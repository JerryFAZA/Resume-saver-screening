# ============================================================
# 智联招聘简历批量下载 — 外部配置文件
# ============================================================

$Config = @{
    # ---- 基础参数（每次任务前必须修改，或通过命令行参数覆盖） ----
    Url            = ""   # 【必填】智联推荐页完整 URL（不含 # 片段），如 "https://rd6.zhaopin.com/app/recommend?tab=recommend&jobNumber=XXX"
    JobName        = ""   # 【必填】岗位名称，需与页面上标签完全一致，如 "销售部经理"
    DownloadDir    = ""   # 【必填】保存目录（绝对路径），如 "C:\Resumes\岗位名"
    DownloadCount  = 5    # 【可选】期望下载数量，默认 5，可调大
    FileFormat     = "word"  # 【可选】期望下载文件格式，"word"（默认，生成 .docx）或 "pdf"。
                              # 注意：文件搜索层同时匹配 .pdf 和 .docx（不依赖此参数），
                              # 以应对网页端记忆了不同格式或用户选错格式的情况（详见 tech_details #24）

    # ---- WebBridge ----
    WebBridgeUrl = "http://127.0.0.1:10086/command"
    Session      = "resume-screening"   # WebBridge 会话标识

    # ---- CDP 鼠标坐标（1920×1080 默认回退值，每次循环优先动态探测） ----
    SaveLocalX   = 1079   # "存至本地"按钮 X（回退值，运行时会动态获取）
    SaveLocalY   = 142    # "存至本地"按钮 Y（回退值，运行时会动态获取）
    SaveConfirmX = 996    # "保存"确认按钮 X（回退值，运行时会动态获取）
    SaveConfirmY = 561    # "保存"确认按钮 Y（回退值，运行时会动态获取）
    MaskCloseX   = 30     # 遮罩关闭区域 X
    MaskCloseY   = 300    # 遮罩关闭区域 Y

    # ---- 时序参数（毫秒） ----
    CloseWaitMs        = 700
    ClickWaitMs        = 800
    MouseMoveWaitMs    = 900
    PressReleaseWaitMs = 400
    DialogCheckWaitMs  = 2000
    DownloadWaitMs     = 6000   # PDF 等待时间；Word 文件通常更大，建议传 10000
    ScrollWaitMs       = 1200

    # ---- 文件管理 ----
    # ★ 注意：脚本内部统一用 Get-ChildItem -Include "*智联简历*.{ext}" 匹配下载文件，
    #   DownloadFilter 仅作调试/文档保留，不再用于 -Filter（中文通配符在 GBK 下会匹配失败）
    DownloadFilter = "*智联简历*"
    DownloadSource = "$env:USERPROFILE\Downloads"

    # ---- 去重规则 ----
    # 判断重复的三个条件（必须同时满足）：
    #   1. 姓名相同（文件名中 _ 分隔的第一部分）
    #   2. 年龄相同（_ 后、岁 前的数字）
    #   3. 文件大小相同（字节数）
    # 发现重复时 → 删除新文件 → 继续下载下一条候选人（不重置流程）
}
