from pywinauto import Application, ElementNotFoundError
import win32gui 
import time  # 导入计时模块

# ---------------- 1. 连接VORTEX Client应用 ----------------
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
        title_re=".*RE小房间-常规.*",  # 模糊匹配子窗口标题
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
    title="e57" 
)
site1_checkbox.wait('visible enabled', timeout=2)
site1_checkbox.click_input()  #点击
print(f"✅ 点击【e57】单选框状态成功")

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
    title="合并" 
)
site1_checkbox.wait('visible enabled', timeout=2)
site1_checkbox.click_input()  #点击
print(f"✅ 点击【合并】单选框状态成功")

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
site_detect_ctrl.set_text("e57+合并+反射率")
print(f"✅ 【文件夹命名】成功")

site_detect_ctrl = Browser_window.child_window(
    control_type="Button",    # 控件类型
    title="确定"        
)
site_detect_ctrl.wait('visible enabled', timeout=2)
site_detect_ctrl.click_input()
print(f"✅ 点击【确定】成功")

# ---------------- 4. 计时功能：监控格式转换过程 ----------------
# 4.1 开始计时（点击文件夹确定后）
start_time = time.time()
start_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(start_time))
print(f"\n⏱️ 格式转换计时开始 - {start_time_str}")

# 4.2 循环检测「格式转换成功」控件/窗口是否出现
timeout = 1000  # 最大超时时间（秒），可根据实际需求调整
check_interval = 0.5  # 检测间隔（0.5秒）
convert_success = False  # 标记是否检测到转换成功

while not convert_success and (time.time() - start_time) < timeout:
    try:
        # -------- 核心修改：定位「格式转换成功」的窗口/文本控件 --------
        success_window = dlg.window(auto_id="MessageForm")
        
        # 检测控件是否可见（出现）
        if success_window.exists() and success_window.is_visible():
            convert_success = True  # 检测到成功，结束循环
            print(f"\n🎉 检测到【格式转换成功】控件/窗口出现！")
        else:
            # 未出现，继续检测，实时输出已耗时
            elapsed = time.time() - start_time
            print(f"🔄 等待格式转换完成... 已耗时：{elapsed:.1f}秒", end="\r")
    except ElementNotFoundError:
        # 控件未找到（还没出现），继续检测
        elapsed = time.time() - start_time
        print(f"🔄 等待格式转换完成... 已耗时：{elapsed:.1f}秒", end="\r")
    
    # 等待检测间隔（仅当未检测到成功时等待）
    if not convert_success:
        time.sleep(check_interval)

# 4.3 结束计时并输出结果
end_time = time.time()
end_time_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(end_time))
total_time = end_time - start_time

# 点击转换成功窗口的【确定】按钮（根据你的实际控件调整）
try:
    site_detect_ctrl = success_window.child_window(
        control_type="Button",    # 控件类型
        title="确定(O)"     
    )
    site_detect_ctrl.wait('visible enabled', timeout=2)
    site_detect_ctrl.click_input()
    print(f"✅ 点击【格式转换成功】的确定按钮成功")
except ElementNotFoundError as e:
    print(f"⚠️ 未找到格式转换成功窗口的确定按钮：{e}")

# 输出最终计时结果
if convert_success:
    print(f"\n✅ 格式转换完成！")
    print(f"📅 开始时间：{start_time_str}")
    print(f"📅 结束时间：{end_time_str}")
    print(f"⏱️ 总耗时：{total_time:.2f}秒（{total_time/60:.2f}分钟）")
else:
    print(f"\n⚠️ 计时超时！最大等待时间{timeout}秒，实际耗时{total_time:.2f}秒")
    print(f"❌ 未检测到【格式转换成功】窗口/控件")