pyinstaller -F better.py
cp .\dist\better.exe .\release\
mv .\release\better.exe .\release\程序.exe
7z a -tzip release.zip .\release\*