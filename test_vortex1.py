from pywinauto import Application, ElementNotFoundError
import win32gui
import time
import psutil
from datetime import datetime
import threading
import csv

# ---------------- 全局变量（资源监控用） ----------------
monitor_data = []  # 存储监控数据
is_monitoring = False  # 监控线程开关
vortex_pid = None  # VORTEX进程PID（用于精准监控）

# ---------------- 资源监控函数（后台线程） ----------------
def monitor_resource(interval=0.3):
    """后台监控系统+VORTEX进程的资源占用"""
    global monitor_data, is_monitoring, vortex_pid
    if not vortex_pid:
        print("⚠️ 未获取到VORTEX进程PID，跳过资源监控")
        return
    
    # 获取VORTEX进程对象
    try:
        vortex_process = psutil.Process(vortex_pid)
    except psutil.NoSuchProcess:
        print("⚠️ VORTEX进程不存在，跳过资源监控")
        return
    
    start_time = time.perf_counter()  # 监控起始时间
    while is_monitoring:
        # 1. 系统级资源
        sys_cpu = psutil.cpu_percent(interval=0)  # 系统CPU使用率(%)
        sys_mem = psutil.virtual_memory().percent  # 系统内存使用率(%)
        
        # 2. VORTEX进程级资源
        try:
            proc_cpu = vortex_process.cpu_percent(interval=0)  # 进程CPU使用率(%)
            proc_mem = vortex_process.memory_percent()  # 进程内存占比(%)
            proc_mem_mb = round(vortex_process.memory_info().rss / 1024 / 1024, 2)  # 进程内存占用(MB)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            proc_cpu, proc_mem, proc_mem_mb = 0, 0, 0
        
        # 3. 记录数据（时间戳+耗时+资源指标）
        elapsed_time = round(time.perf_counter() - start_time, 3)
        monitor_data.append({
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],  # 毫秒级时间戳
            "elapsed_time(s)": elapsed_time,
            "sys_cpu(%)": sys_cpu,
            "sys_mem(%)": sys_mem,
            "vortex_cpu(%)": proc_cpu,
            "vortex_mem(%)": proc_mem,
            "vortex_mem(MB)": proc_mem_mb
        })
        time.sleep(interval)

# ---------------- 保存监控数据到CSV ----------------
def save_monitor_data():
    """将监控数据保存为CSV文件（按时间命名）"""
    global monitor_data
    if not monitor_data:
        print("⚠️ 无监控数据可保存")
        return
    
    # 生成带时间戳的文件名
    filename = f"精细识别监控报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    # 写入CSV
    with open(filename, "w", newline="", encoding="utf-8") as f:
        fieldnames = ["timestamp", "elapsed_time(s)", "sys_cpu(%)", "sys_mem(%)",
                     "vortex_cpu(%)", "vortex_mem(%)", "vortex_mem(MB)"]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(monitor_data)
    print(f"✅ 监控数据已保存至：{filename}")

# ---------------- 核心业务逻辑 ----------------
if __name__ == "__main__":
    # ---------------- 1. 句柄连接VORTEX主窗口 ----------------
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
            title_re=".*3-0.6-(2)/站点识别.*",  # 模糊匹配子窗口标题
            control_type="Window",             # 限定为窗口类型
            parent=dlg                         # 限定父窗口
        )
        vortex_window.wait('visible', timeout=5)
        print(f"✅ 定位到目标窗口：标题={vortex_window.window_text()}，句柄={hex(vortex_window.handle)}")
    except ElementNotFoundError as e:
        print(f"❌ 目标子窗口定位失败：{e}")
        exit(1)

    # ---------------- 3. 标靶识别功能操作 + 监控 ----------------
    try:
        # 3.1 定位并点击【站点识别】
        site_detect_ctrl = dlg.child_window(
            auto_id="btnDetect",    # 控件唯一标识
            control_type="Pane",    # 控件类型
            title="站点识别"        
        )
        site_detect_ctrl.wait('visible enabled', timeout=5)
        site_detect_ctrl.click_input()
        print(f"✅ 点击【站点识别】成功")

        # 3.2 定位并切换【站点1】复选框
        site1_checkbox = dlg.child_window(
            control_type="CheckBox", # 复选框类型
            title="站点1"             # ✅ 修正：UIA控件用name而非title
        )
        site1_checkbox.wait('visible enabled', timeout=5)
        site1_checkbox.toggle()  # 切换勾选状态
        print(f"✅ 切换【站点1】复选框状态成功")

        # ---------------- 4. 【精细识别】操作 + 耗时/资源监控 ----------------
        # 4.1 启动后台资源监控
        is_monitoring = True
        monitor_thread = threading.Thread(target=monitor_resource, daemon=True)
        monitor_thread.start()
        print("📊 开始监控资源占用...")


        # 4.3 定位并操作【精细识别】
        fine_detect_ctrl = dlg.child_window(
            control_type="Pane",    # 控件类型
            title="精细识别"         # ✅ 修正：UIA控件用name而非title
        )
        fine_detect_ctrl.wait('visible enabled', timeout=5)
        fine_detect_ctrl.click_input()  # 执行精细识别操作
        print(f"✅ 点击【精细识别】成功，等待操作完成（监控标靶编辑按钮状态）...")

        # 4.2 记录【精细识别】操作开始时间
        fine_detect_start = time.perf_counter()

        # ---------------- 关键修改：等待【标靶编辑】按钮可点击（判断精细识别结束） ----------------
        # 定位【标靶编辑】控件（用AutoID最精准，避免重名）
        target_edit_ctrl = dlg.child_window(
            auto_id="btn_edit",     # 唯一标识（优先用这个，比name更稳定）
            control_type="Pane",    # 控件类型：UIA_PaneControlTypeId
            title="标靶编辑"         # 双重验证，确保定位正确
        )
        # 等待控件变为【可点击状态（IsEnabled=True）+ 可见】，超时30秒（可根据实际调整）
        target_edit_ctrl.wait(
            'visible enabled',     # 等待条件：可见且可点击
            timeout=10000             # 最大等待时间（秒），防止卡死
        )
        print(f"✅ 【标靶编辑】按钮已可点击，精细识别操作完成！")

        # 4.5 记录操作结束时间，计算耗时
        fine_detect_end = time.perf_counter()
        fine_detect_duration = round(fine_detect_end - fine_detect_start, 3)
        print(f"⏱️ 【精细识别】操作总耗时：{fine_detect_duration} 秒")

    except ElementNotFoundError as e:
        print(f"❌ 控件定位失败：{e}")
    except Exception as e:
        print(f"❌ 操作异常：{e}")
    finally:
        # 停止监控并保存数据
        is_monitoring = False
        if 'monitor_thread' in locals():
            monitor_thread.join(timeout=2)  # 等待监控线程结束
        save_monitor_data()
        print("🔚 监控结束")