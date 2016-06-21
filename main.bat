@echo off
cd "C:\Users\Shitikantha\Desktop\EngWord_in_odia_wikt"

del out.txt
rem --to avoiding the over-writting of data we delete the existing data
echo.
e2oRead.py
echo done.
echo.
echo output stored at 'out.txt'
pause