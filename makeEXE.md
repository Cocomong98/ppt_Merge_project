- venv 설치
pip install pyinstaller pyqt5 python-pptx pywin32

- exe파일 생성 


ORIGIN
pip install pyinstaller pyqt5 python-pptx pywin32
4.  **EXE 생성 (Windows):** 다음 명령어를 실행하여 `.exe` 파일을 생성합니다.
```bash
pyinstaller --name "PPT_Merger" --onefile -w \
    --hidden-import "win32com" \
    --collect-all "pyqt5" \
    --collect-all "win32com" \
    app.py

LITE
pyinstaller --name "PPT_Merger" --onefile -w --hidden-import "win32com" --collect-all "pyqt5" --collect-all "win32com" app.py