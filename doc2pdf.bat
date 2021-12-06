@echo off
setlocal

set input=%1
if "%input%"=="" set input=*.docx

for %%a in (%input%) do powershell -f doc2pdf.ps1 "%%a" "%%~dpna.pdf"
)
endlocal
