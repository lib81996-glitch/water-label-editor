@echo off
chcp 936 >nul
setlocal EnableDelayedExpansion
title 水洗标签编辑器 - 运行环境检测

REM ===================================================================
REM  水洗标签编辑器 · 运行环境检测
REM  用法：双击运行             → 检测并给出下载菜单
REM        运行环境检测.bat /auto → 只检测，不等按键（供自动验证用）
REM  注意：只检测、只在手动选 1 时下载；绝不静默安装任何东西。
REM ===================================================================

set "HERE=%~dp0"
set "AUTO=%~1"
set "FAIL=0"
set "WARN=0"

echo.
echo ============================================================
echo    水洗标签编辑器   ·   运行环境检测
echo ============================================================
echo.
echo   软件目录：%HERE%
echo   本工具只做检查；下载需你手动选择，绝不自动安装。
echo.

REM ---------- 1. 系统与位数 ----------
echo [1/7] 操作系统
for /f "delims=" %%i in ('ver') do echo        %%i
echo       处理器架构：%PROCESSOR_ARCHITECTURE%
if /i "%PROCESSOR_ARCHITECTURE%"=="AMD64" goto arch_ok
echo       [提示] 不是 64 位环境，请确认 exe 能启动。
set /a WARN+=1
goto arch_done
:arch_ok
echo       [正常] 64 位系统
:arch_done
echo.

REM ---------- 2. 中文字体（缺了标签会打成空白）----------
echo [2/7] 中文字体
set "FONT=0"
if exist "%WINDIR%\Fonts\msyh.ttc"  set "FONT=1"
if exist "%WINDIR%\Fonts\simsun.ttc" set "FONT=1"
if exist "%WINDIR%\Fonts\simhei.ttf" set "FONT=1"
if "%FONT%"=="1" goto font_ok
echo       [缺失] 没找到微软雅黑 / 宋体 / 黑体。
echo              标签会打成空白（不是方块，是一个字都没有）。
echo              解决：设置 - 时间和语言 - 语言 - 添加中文语言包，
echo              或从另一台中文电脑拷贝 msyh.ttc 到 Fonts 文件夹。
set /a FAIL+=1
goto font_done
:font_ok
echo       [正常] 找到了中文字体
:font_done
echo.

REM ---------- 3. 文件夹可写（保存 / 自动恢复 / 复制图标都靠它）----------
echo [3/7] 软件文件夹是否可写
set "TESTF=%HERE%._envtest.tmp"
echo ok>"%TESTF%" 2>nul
if not exist "%TESTF%" goto dir_bad
del "%TESTF%" >nul 2>&1
echo       [正常] 可以写入
goto dir_done
:dir_bad
echo       [缺失] 这个文件夹不能写，保存、自动恢复、复制图标都会失败。
echo              解决：把整个文件夹挪到桌面或 D/E 盘，别放 Program Files，
echo              也别放在带「只读」属性的 U 盘里。
set /a FAIL+=1
:dir_done
echo.

REM ---------- 4. 运行库 ----------
echo [4/7] VC++ 运行库
if exist "%WINDIR%\System32\vcruntime140.dll" goto vc_ok
echo       [缺失] 没找到 vcruntime140.dll，程序可能双击没反应。
echo              解决：装「Microsoft Visual C++ 2015-2022 运行库(x64)」，
echo              菜单按 1 可从微软官网下载。
set /a FAIL+=1
set "NEEDVC=1"
goto vc_done
:vc_ok
echo       [正常] vcruntime140.dll 存在
set "NEEDVC=0"
:vc_done
echo.

REM ---------- 5. 软件文件是否齐全 ----------
echo [5/7] 软件文件完整性
set "MISS=0"
for %%F in (
  "水洗标签编辑器.exe"
  "assets\care_block.png"
  "assets\top_rule_ref.png"
  "assets\icons\icon1.png"
  "assets\icons\icon2.png"
  "assets\icons\icon3.png"
  "assets\icons\icon4.png"
  "assets\icons\icon5.png"
  "assets\icons\source_strip.png"
  "标准底版-35x75.wslabel"
  "示例-6608.wslabel"
  "使用说明.txt"
) do (
  if not exist "%HERE%%%~F" (
    echo       [缺文件] %%~F
    set /a MISS+=1
  )
)
if "%MISS%"=="0" goto files_ok
echo              共缺 %MISS% 个文件，缺图标会导致洗涤说明那行空白。
echo              解决：重新完整解压压缩包，不要只拷 exe 一个文件。
set /a FAIL+=1
goto files_done
:files_ok
echo       [正常] 12 个必需文件齐全
:files_done
echo.

REM ---------- 6. 打印机 ----------
REM ★ 不用 Get-Printer：它走 WMI 的 ROOT/StandardCimv2，有些机器这个
REM   命名空间坏了会报「无效命名空间」，明明装了打印机也读成 0 台。
REM   改读注册表 HKLM:\...\Print\Printers 的键，绕开 WMI，实测能列出
REM   Gprinter 条码机。Spooler 服务顺带查一下，没起也打不了。
REM ★ for /f in ('...') 用单引号界定命令，命令内部**绝不能**再出现单引号，
REM   否则界定串被提前关闭、剩下的字被 CMD 当命令执行（报「此时不应有 x。」）。
echo [6/7] 打印机
set "PRN=0"
for /f "delims=" %%p in ('powershell -NoProfile -Command "try { @(Get-ChildItem HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers -ErrorAction Stop).Count } catch { 0 }" 2^>nul') do set "PRN=%%p"
if "%PRN%"=="0" goto prn_bad
echo       [正常] 注册表里有 !PRN! 台打印机
for /f "delims=" %%s in ('powershell -NoProfile -Command "try { (Get-Service Spooler -ErrorAction Stop).Status } catch { Unknown }" 2^>nul') do set "SPOOL=%%s"
if /i "!SPOOL!"=="Running" goto spool_ok
echo       [提示] 后台打印服务 Spooler 没在运行，打印会失败。
echo              解决：服务里把 Print Spooler 设为自动并启动。
set /a WARN+=1
goto prn_done
:spool_ok
echo       [正常] Print Spooler 正在运行
goto prn_done
:prn_bad
echo       [提示] 没读到打印机（可能是没装驱动，或权限限制）。
echo              打印前请在 设置 - 打印机 里确认条码机在线。
set /a WARN+=1
:prn_done
echo.

REM ---------- 7. 屏幕分辨率（窗口超屏风险）----------
REM 同上：用 [char]120 代替 'x' 字符串，避免单引号破坏 for /f 的界定。
echo [7/7] 屏幕分辨率
set "RES="
for /f "delims=" %%r in ('powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $b=[System.Windows.Forms.Screen]::PrimaryScreen.Bounds; $b.Width.ToString() + [char]120 + $b.Height.ToString()" 2^>nul') do set "RES=%%r"
if defined RES goto res_ok
echo       [提示] 没读到分辨率，不影响使用。
goto res_done
:res_ok
echo       主屏：!RES!
echo              若分辨率宽高小于 1200 x 800 或缩放设为 125%%/150%% ，
echo              软件启动时会自动把窗口收进屏幕，并可拖四边调整大小。
:res_done
echo.

REM ---------- 结论 ----------
echo ============================================================
if not "%FAIL%"=="0" goto sum_bad
echo   检测完成：没有发现阻断性问题，可以直接双击 水洗标签编辑器.exe
if not "%WARN%"=="0" echo   另有 %WARN% 项提示，不影响启动，按需处理。
goto sum_done
:sum_bad
echo   检测完成：发现 %FAIL% 项需要处理，处理前软件可能打不开或打空白。
:sum_done
echo ============================================================
echo.
echo   如需下载缺失组件，可用下面的菜单（来源均为官方）：
echo.
echo     [1] Microsoft Visual C++ 2015-2022 运行库 x64
echo         官方地址：https://aka.ms/vs/17/release/vc_redist.x64.exe
echo     [2] 查看中文语言包安装指引（不下载文件）
echo     [0] 退出
echo.
echo   下载的文件会放在本文件夹的 _下载 子目录，需要你手动双击安装。
echo.

if /i "%AUTO%"=="/auto" exit /b 0

set "CH="
set /p "CH=  请输入数字后回车："
if "%CH%"=="1" goto dl_vc
if "%CH%"=="2" goto font_help
goto bye

:dl_vc
set "DL=%HERE%_下载"
if not exist "%DL%" mkdir "%DL%" >nul 2>&1
echo.
echo   正在从微软官方下载，请稍候（视网络可能需要一两分钟）...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -Uri 'https://aka.ms/vs/17/release/vc_redist.x64.exe' -OutFile '%DL%\vc_redist.x64.exe' -UseBasicParsing; 'OK' } catch { 'FAIL' }"
if exist "%DL%\vc_redist.x64.exe" goto dl_ok
echo   下载没成功。请检查网络，或手动打开上面那个官方地址下载。
goto bye
:dl_ok
echo.
echo   已下载：%DL%\vc_redist.x64.exe
echo   请手动双击它完成安装，然后重新运行本检测。
goto bye

:font_help
echo.
echo   安装中文语言包：Windows 设置 - 时间和语言 - 语言 -
echo   「添加语言」选 中文(简体，中国) - 安装后重启。
echo   或者从一台中文 Windows 的 C:\Windows\Fonts 拷贝
echo   msyh.ttc 到本机同目录（需管理员权限）。
goto bye

:bye
echo.
pause
exit /b 0
