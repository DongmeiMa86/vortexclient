from pywinauto import Application, ElementNotFoundError
import win32gui
import time
import psutil
from datetime import datetime
import threading
import csv

# ---------------- 1. 保留句柄连接主窗口（你的核心需求） ----------------
hwnd = win32gui.FindWindow(None, "VORTEX Client")
if not hwnd:
    raise Exception("❌ 未找到标题为'VORTEX Client'的窗口，请确认程序已启动！")

# 连接主窗口（UIA后端）
app = Application(backend="uia").connect(handle=hwnd)
dlg = app.window(handle=hwnd)
dlg.wait('visible', timeout=10)
dlg.set_focus()
print(f"✅ 成功连接VORTEX主窗口：句柄={hex(hwnd)}，标题={dlg.title}")

# ---------------- 2. 定位目标子窗口（放弃class_name，改用唯一属性） ----------------
# 方案1：按名称定位「标靶识别与编辑」相关窗口（优先推荐）
# （替换为你实际的目标窗口名称，比如"pro103，8/站点识别"）

    # 定位顶层子窗口（top_level_only=True 仅匹配顶层窗口，避免匹配子控件）
vortex_window = dlg.window(
    title_re=".*pro103，8/站点识别.*",  # 模糊匹配名称（正则）
    control_type="Window",         # 限定为窗口类型
    parent = dlg                  # 限定父窗口为VORTEX主窗口
)
vortex_window.wait('visible', timeout=5)
print(f"✅ 定位到目标窗口：标题={vortex_window.title}，句柄={hex(vortex_window.handle)}")

#---------------- 3. 标靶识别功能控件定位操作 ----------------
try:
    detect_ctrl = dlg.child_window(
        auto_id="btnDetect",    # 唯一标识，确保不匹配其他控件
        control_type="Pane",     # 限定类型，缩小范围
        title = "站点识别"
    )
    detect_ctrl.wait('visible', timeout=5)
    detect_ctrl.click_input()
    print(f"✅ 定位到站点识别")

    detect_ctrl = dlg.child_window(
        control_type="CheckBox",     # 限定类型，缩小范围
        title = "站点1"
    )
    detect_ctrl.wait('visible', timeout=5)
    detect_ctrl.toggle()
    print(f"✅ 定位到站点1")

    detect_ctrl = dlg.child_window(
        control_type="Pane",     # 限定类型，缩小范围
        title = "精细识别"
    )
    detect_ctrl.wait('visible', timeout=5)
    detect_ctrl.toggle()
    print(f"✅ 定位到精细识别")
except ElementNotFoundError as e:
     print(f"未找到指定的控件: {e}")

