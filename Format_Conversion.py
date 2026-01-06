from pywinauto import Application, ElementNotFoundError
import win32gui 

#连接VORTEX Client应用
hwnd = win32gui.FindWindow(None, "VORTEX Client")
if not hwnd:
    raise Exception("❌ 未找到标题为'VORTEX Client'的窗口，请确认程序已启动！")

# 连接主窗口（UIA后端）
app = Application(backend="uia").connect(handle=hwnd)
dlg = app.window(handle=hwnd)
dlg.wait('visible enabled', timeout=10)
dlg.set_focus()
vortex_pid = dlg.process_id()  # 获取VORTEX进程PID（关键！）
print(f"✅ 成功连接VORTEX主窗口：句柄={hex(hwnd)}，标题={dlg.window_text()}，PID={vortex_pid}")

# ---------------- 2. 定位目标子窗口 ----------------
try:
    vortex_window = dlg.window(
        title_re=".*pro103，8.*",  # 模糊匹配子窗口标题
        control_type="Window"            # 限定为窗口类型
    )
    vortex_window.wait('visible', timeout=5)
    print(f"✅ 定位到目标窗口：标题={vortex_window.window_text()}，句柄={hex(vortex_window.handle)}")
except ElementNotFoundError as e:
    print(f"❌ 目标子窗口定位失败：{e}")
    exit(1)

# ---------------- 3. 格式转换核心操作 ----------------
# 3.1 定位并点击【导出】
site_detect_ctrl = vortex_window.child_window(
    control_type="Button",    # 控件类型
    title="导出"        
)
site_detect_ctrl.wait('visible enabled', timeout=5)
site_detect_ctrl.click_input()
print(f"✅ 点击【导出】成功")

# 3.2 定位并点击【点云】
option_window = dlg.window(
        title_re=".*选项.*",  # 模糊匹配子窗口标题
        control_type="Window"            # 限定为窗口类型
    )
option_window.wait('visible', timeout=5)
print(f"✅ 定位到目标窗口：标题={option_window.window_text()}，句柄={hex(option_window.handle)}")

site_detect_ctrl = option_window.child_window(
    control_type="Pane",    # 控件类型
    title="点云"        
)
site_detect_ctrl.wait('visible enabled', timeout=2)
site_detect_ctrl.click_input()
print(f"✅ 点击【点云】成功")

# 3.3 选择输出格式
export_window = dlg.window(
        title_re=".*点云导出.*",  # 模糊匹配子窗口标题
        control_type="Window"            # 限定为窗口类型
    )
export_window.wait('visible', timeout=5)
print(f"✅ 定位到目标窗口：标题={export_window.window_text()}，句柄={hex(export_window.handle)}")

site1_checkbox = export_window.child_window(
    control_type="RadioButton", # 单选按钮类型
    title="pts" 
)
site1_checkbox.wait('visible enabled', timeout=2)
site1_checkbox.click_input()  #点击
print(f"✅ 点击【pts】单选框状态成功")

# 3.4 是否启用点云抽稀
site1_checkbox = export_window.child_window(
    control_type="CheckBox", # 复选框类型
    title="启用" 
)
site1_checkbox.wait('visible enabled', timeout=2)
site1_checkbox.toggle()  #切换状态
print(f"✅ 切换【启用】复选框状态成功")

# 3.5 选择输出类型
site1_checkbox = export_window.child_window(
    control_type="RadioButton", # 单选按钮类型
    title="单站" 
)
site1_checkbox.wait('visible enabled', timeout=2)
site1_checkbox.click_input()  #点击
print(f"✅ 点击【单站】单选框状态成功")

# 3.6 选择贴图
site1_checkbox = export_window.child_window(
    control_type="RadioButton", # 单选按钮类型
    title="反射率" 
)
site1_checkbox.wait('visible enabled', timeout=2)
site1_checkbox.click_input()  #点击
print(f"✅ 点击【反射率】单选框状态成功")

# 3.7 是否点云降噪
site1_checkbox = export_window.child_window(
    control_type="CheckBox", # 复选框类型
    title="点云降噪" 
)
site1_checkbox.wait('visible enabled', timeout=2)
site1_checkbox.toggle()  #切换状态
print(f"✅ 切换【点云降噪】复选框状态成功")

# 3.8 是否点云厚度优化
site1_checkbox = export_window.child_window(
    control_type="CheckBox", # 复选框类型
    title="点云厚度优化" 
)
site1_checkbox.wait('visible enabled', timeout=2)
site1_checkbox.toggle()  #切换状态
print(f"✅ 切换【点云厚度优化】复选框状态成功")

# 3.9 定位并点击【导出】
site_detect_ctrl = export_window.child_window(
    control_type="Pane",    # 控件类型
    title="导出",
    auto_id="uiButton3"        
)
site_detect_ctrl.wait('visible enabled', timeout=2)
site_detect_ctrl.click_input()
print(f"✅ 点击【导出】成功")

# 3.10 选择文件输出路径（此电脑→F盘→）
Browser_window = dlg.window(
        title_re=".*浏览文件夹.*",  # 模糊匹配子窗口标题
        control_type="Window"            # 限定为窗口类型
    )
Browser_window.wait('visible', timeout=5)
print(f"✅ 定位到目标窗口：标题={Browser_window.window_text()}，句柄={hex(Browser_window.handle)}")

site_detect_ctrl = Browser_window.child_window(
    control_type="TreeItem",    # 控件类型
    title="此电脑"     
)
site_detect_ctrl.wait('visible enabled', timeout=2)
site_detect_ctrl.click_input()
print(f"✅ 点击【此电脑】成功")

site_detect_ctrl = Browser_window.child_window(
    control_type="TreeItem",    # 控件类型
    title="新加卷 (F:)"     
)
site_detect_ctrl.wait('visible enabled', timeout=2)
site_detect_ctrl.click_input()
print(f"✅ 点击【F盘】成功")

site_detect_ctrl = Browser_window.child_window(
    control_type="Button",    # 控件类型
    title="新建文件夹(M)"     
)
site_detect_ctrl.wait('visible enabled', timeout=2)
site_detect_ctrl.click_input()
print(f"✅ 点击【新建文件夹】成功")

site_detect_ctrl = Browser_window.child_window(
    control_type="Edit",    # 控件类型
    auto_id="1"     
)
site_detect_ctrl.wait('visible enabled', timeout=7)
site_detect_ctrl.set_text("pts+单站+反射率")
print(f"✅ 【文件夹命名】成功")

site_detect_ctrl = Browser_window.child_window(
    control_type="Button",    # 控件类型
    title="确定"        
)
site_detect_ctrl.wait('visible enabled', timeout=2)
site_detect_ctrl.click_input()
print(f"✅ 点击【确定】成功")