@echo off
chcp 65001 >nul
cd /d %~dp0

echo ========================================
echo   中国银行信用卡账单处理流水线
echo ========================================
echo.

REM ============================================================
REM 步骤 1：检查 Python
REM ============================================================
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python
    echo        请将 Python 添加到系统 PATH，或重启命令行窗口
    pause
    exit /b 1
)

REM ============================================================
REM 步骤 2：创建虚拟环境（如果不存在）
REM ============================================================
if not exist .venv\Scripts\python.exe (
    echo [初始化] 首次运行，正在创建虚拟环境...
    python -m venv .venv
    if errorlevel 1 (
        echo [错误] 创建虚拟环境失败
        pause
        exit /b 1
    )
)

REM ============================================================
REM 步骤 3：检查并安装依赖
REM   - 关键修复：不只看 .venv 是否存在，要看 .deps_installed 标记
REM   - 标记文件被删除/缺失 → 重新装依赖
REM ============================================================
call .venv\Scripts\activate

REM 用一个轻量探测判断关键依赖是否齐备
.venv\Scripts\python.exe -c "import pandas, pdfplumber, openpyxl, matplotlib, dotenv" >nul 2>&1
if errorlevel 1 (
    echo.
    echo ============================================================
    echo  [初始化] 正在安装依赖（首次约需 1-3 分钟）
    echo           使用清华镜像源加速下载
    echo           请耐心等待，不要关闭此窗口
    echo ============================================================
    echo.

    REM 配置 pip 使用清华镜像源（避免超时）
    .venv\Scripts\python.exe -m pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple
    .venv\Scripts\python.exe -m pip config set global.trusted-host pypi.tuna.tsinghua.edu.cn

    REM 升级 pip（避免旧版 pip 的解析问题）
    .venv\Scripts\python.exe -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple

    REM 安装依赖（指定超时时间，给出详细输出）
    .venv\Scripts\python.exe -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --timeout 300
    if errorlevel 1 (
        echo.
        echo [错误] 依赖安装失败
        echo        常见原因：网络问题、镜像源不可用
        echo        可尝试手动执行：
        echo          .venv\Scripts\activate
        echo          pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
        pause
        exit /b 1
    )

    REM 创建标记文件：表示依赖已完整安装
    echo Installed at %date% %time% > .deps_installed
    echo.
    echo [OK] 依赖安装完成
    echo.
) else (
    REM 二次确认：标记文件丢失也强制重装
    if not exist .deps_installed (
        echo [初始化] 依赖标记缺失，重新安装...
        .venv\Scripts\python.exe -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --timeout 300
        if errorlevel 1 (
            echo [错误] 依赖安装失败
            pause
            exit /b 1
        )
        echo Installed at %date% %time% > .deps_installed
    )
)

REM ============================================================
REM 步骤 4：检查 .env
REM ============================================================
if not exist .env (
    echo.
    echo [提示] 未找到 .env 文件，自动下载账单将被跳过
    echo       如需自动下载，请执行：
    echo         copy .env.example .env
    echo       然后编辑 .env 填入邮箱账号和授权码
    echo.
)

echo.
echo [开始] 启动流水线（交互模式：在人工介入点会暂停等待）...
echo.

python -m bill_pipeline.cli all --interactive

echo.
echo ========================================
echo   流水线结束
echo ========================================
pause
