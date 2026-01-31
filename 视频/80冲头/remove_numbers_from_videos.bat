@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo 开始重命名视频文件...
echo.

set "count=0"

for %%f in (*.mp4 *.avi *.mov *.mkv *.flv *.wmv *.webm *.m4v) do (
    set "filename=%%~nf"
    set "ext=%%~xf"
    set "original=%%~nxf"
    
    call :remove_trailing_numbers "!filename!" newname
    
    if not "!newname!"=="" (
        if not "!newname!"=="!filename!" (
            set "newfile=!newname!!ext!"
            
            if exist "!newfile!" (
                echo [跳过] "!original!" - 目标文件已存在: "!newfile!"
            ) else (
                ren "!original!" "!newfile!"
                echo [重命名] "!original!" -^> "!newfile!"
                set /a count+=1
            )
        )
    )
)

echo.
echo 完成！共重命名 !count! 个文件。
pause
goto :eof

:remove_trailing_numbers
set "input=%~1"
:loop
set "lastchar=%input:~-1%"
echo %lastchar%| findstr /r "[0-9]" >nul
if %errorlevel% equ 0 (
    set "input=%input:~0,-1%"
    goto loop
)
set "%~2=%input%"
goto :eof
