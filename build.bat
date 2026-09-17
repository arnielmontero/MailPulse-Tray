pyinstaller --noconsole --onefile --name MailPulseTray ^
  --hidden-import=win32timezone ^
  --hidden-import=win32com ^
  --hidden-import=win32com.client ^
  --hidden-import=pythoncom ^
  --hidden-import=pywintypes ^
  --collect-submodules win32com ^
  main.py
