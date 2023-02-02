# 说明

**使用如下额外库：**

```python
import openpyxl
import hashlib
from PySide2.QtGui import QIcon
from PySide2.QtWidgets import QApplication, QMessageBox, QFileDialog, QTableWidgetItem, QLineEdit, QInputDialog
from PySide2.QtUiTools import QUiLoader
from configobj import ConfigObj
import subprocess
import os
import pyperclip
import keyboard
```

**自动P图片文字，自动发送彩信，循环式自动电脑程序**

# 程序打包命令

pyinstaller main.py --noconsole --uac-admin --hidden-import PySide2.QtXml --icon="logo.ico"