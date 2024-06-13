pyinstaller ^
--onedir ^
--upx-dir "Z:\Development\upx" ^
--hidden-import "babel.numbers" ^
--hidden-import "windows_toasts" ^
--hidden-import "winrt.windows.foundation.collections" ^
--uac-admin ^
--name jars ^
main.py