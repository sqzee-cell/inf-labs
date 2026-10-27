@echo off
setlocal EnableExtensions DisableDelayedExpansion
rem Variant 6. All generated files stay beside this script, under work.
cd /d "%~dp0" || exit /b 1
set "WORK_ROOT=work"
set /a RUN_ID=0 >nul
:choose_working_directory
if not exist "%WORK_ROOT%\Name1" goto start_demo
set /a RUN_ID+=1 >nul
set "WORK_ROOT=work\run_%RUN_ID%"
goto choose_working_directory
:start_demo
if not exist "source.txt" exit /b 3
md "%WORK_ROOT%\Name1\Name2\Name3" || exit /b 4
md "%WORK_ROOT%\Name1\Name2\Name4" || exit /b 4
md "%WORK_ROOT%\Name1\Name5" || exit /b 4
cd /d "%WORK_ROOT%" || exit /b 4
echo === Creating Test1.txt with COPY ===
copy /b "%~dp0source.txt" "Name1\Test1.txt" || exit /b 5
type "Name1\Test1.txt"
echo === Copying the source to all other directories ===
copy /b "Name1\Test1.txt" "Name1\Name2\Test1.txt" || exit /b 5
copy /b "Name1\Test1.txt" "Name1\Name2\Name3\Test1.txt" || exit /b 5
copy /b "Name1\Test1.txt" "Name1\Name2\Name4\Test1.txt" || exit /b 5
copy /b "Name1\Test1.txt" "Name1\Name5\Test1.txt" || exit /b 5
echo === Renaming three copies ===
ren "Name1\Name2\Test1.txt" "Test2.txt" || exit /b 6
ren "Name1\Name2\Name3\Test1.txt" "Test3.txt" || exit /b 6
ren "Name1\Name2\Name4\Test1.txt" "Test4.txt" || exit /b 6
echo === Concatenating Test1.txt and Test2.txt ===
copy /b "Name1\Test1.txt"+"Name1\Name2\Test2.txt" "Name1\Test5.txt" || exit /b 7
type "Name1\Test5.txt"
echo === Directory contents ===
dir "Name1"
echo === Directory tree in the working area ===
tree "Name1" /f /a
echo Completed successfully.
exit /b 0
