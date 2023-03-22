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
**该程序类似电脑版的按键精灵，不过因为很多人不会编程，面对不同的频繁操作不知道如何书写代码，可以使用这个程序进行简单的电脑自动操作**
**内置了开源的QTScrcpy投屏工具，可以通过电脑控制手机**

# 程序打包命令

pyinstaller main.py --noconsole --uac-admin --hidden-import PySide2.QtXml --icon="logo.ico"

# 程序预览图
![](https://github.com/AJAskr/Ps-SendSSM/blob/master/%E9%A2%84%E8%A7%88%E5%9B%BE.png?raw=true)
