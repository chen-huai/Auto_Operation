import pyautogui
import time


class Windows_auto():
    def get_desktop_coordinates(self, second):
        time.sleep(second)
        # x屏幕横轴，y屏幕纵轴
        x, y = pyautogui.position()
        return (x, y)
    def window_automation_operation(self, step):
        # step是个列表
        flag = 1
        try:
            if step[0].lower() == 'move':
                pyautogui.moveTo(int(step[1]), int(step[2]), duration=float(step[3]))
            elif step[0].lower() == 'click':
                pyautogui.click(int(step[1]), int(step[2]))
            elif step[0].lower() == 'dragto':
                pyautogui.dragTo(int(step[1]), int(step[2]), duration=float(step[3]))
            elif step[0].lower() == 'mousedown':
                pyautogui.mouseDown(int(step[1]), int(step[2]), duration=step[3])
            elif step[0].lower() == 'mouseup':
                pyautogui.mouseUp(int(step[1]), int(step[2]), duration=step[3])
            elif step[0].lower() == 'typewrite':
                pyautogui.typewrite(step[1])
            elif step[0].lower() == 'sleep':
                time.sleep(int(step[1]))
            elif step[0].lower() == 'press':
                pyautogui.press(step[1])
            elif step[0].lower() == 'hotkey':
                i = 0
                buttons = []
                for button in step:
                    if i != 0:
                        buttons.append(button)

                    i += 1
                key_string = ', '.join([f"'{key}'" for key in buttons])
                a = pyautogui.hotkey(key_string)
                print(a)
            else:
                time.sleep(3)
        except:
            flag = 0
        return flag
