pyinstaller ^
--onedir ^
--upx-dir "Z:\Development\upx" ^
--hidden-import "babel.numbers" ^
--hidden-import "windows_toasts" ^
--hidden-import "winrt.windows.foundation.collections" ^
--hidden-import "pyfiglet.fonts" ^
--collect-data "pyfiglet.fonts" ^
--uac-admin ^
--name jars ^
main.py