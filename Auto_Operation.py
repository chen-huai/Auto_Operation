import pyautogui

width, height = pyautogui.size()
print(width, height)

# pyautogui.moveTo(0, 0, duration=0.0025)

# pyautogui.click
# pyautogui.FAILSAFE

# pyautogui.moveTo(1919, 1079, duration=0.0025)
img = pyautogui.screenshot()
# 获取当前坐标
x, y = pyautogui.position()
pyautogui.moveTo(1200, 890, duration=0.25)
pyautogui.moveTo(1300, 122, duration=0.25)



from pywinauto import application

# 启动应用程序
app = application.Application().start("C:\\Program Files\\Notepad3\\Notepad3.exe")

# 选择窗口
window = app.UntitledNotepad

# 获取窗口中的所有子控件
children = window.children()

# 输出所有子控件的类名和标题
for child in children:
    print(child.class_name(), child.window_text(),1)
    buttons = child.children()
    for button in buttons:
        print(button,2)
