from pywinauto import Application, ElementNotFoundError
import win32gui 
import time
import json
import os
from datetime import datetime
import logging

# ==================== 配置部分 ====================
class PointCloudConversionTest:
    def __init__(self):
        # 配置日志
        self.setup_logging()
        
        # 测试配置
        self.test_config = {
            'output_format': 'e57',
            'export_type': '合并',
            'texture': '反射率',
            'output_path': 'F:\\e57+合并+反射率',
            'timeout': 1000,
            'check_interval': 0.5
        }
        
        # 测试结果
        self.test_results = {
            'test_name': '点云格式转换自动化测试',
            'start_time': None,
            'end_time': None,
            'duration': None,
            'status': '未执行',
            'steps': [],
            'errors': []
        }

    def setup_logging(self):
        """配置日志系统"""
        log_dir = "logs"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_dir, f"pointcloud_test_{timestamp}.log")
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

# ==================== 测试用例执行 ====================
class PointCloudConversionTestCase:
    def __init__(self, test_manager):
        self.tm = test_manager
        self.logger = test_manager.logger
        self.app = None
        self.dlg = None
        
    def setup(self):
        """测试前置条件设置"""
        self.logger.info("="*60)
        self.logger.info("开始执行点云格式转换自动化测试")
        self.logger.info(f"测试配置: {json.dumps(self.tm.test_config, indent=2, ensure_ascii=False)}")
        
        # 记录开始时间
        self.tm.test_results['start_time'] = datetime.now().isoformat()
        
    def test_step_1_connect_to_vortex(self):
        """步骤1: 连接到VORTEX Client应用"""
        self.logger.info("\n--- 步骤1: 连接VORTEX Client应用 ---")
        try:
            hwnd = win32gui.FindWindow(None, "VORTEX Client")
            if not hwnd:
                raise Exception("未找到标题为'VORTEX Client'的窗口")
            
            self.app = Application(backend="uia").connect(handle=hwnd)
            self.dlg = self.app.window(handle=hwnd)
            self.dlg.wait('visible enabled', timeout=10)
            self.dlg.set_focus()
            
            vortex_pid = self.dlg.process_id()
            self.logger.info(f"✅ 成功连接VORTEX主窗口: 句柄={hex(hwnd)}, PID={vortex_pid}")
            
            self.tm.test_results['steps'].append({
                'step': 1,
                'description': '连接VORTEX Client应用',
                'status': 'PASS',
                'details': f"句柄: {hex(hwnd)}, PID: {vortex_pid}"
            })
            return True
            
        except Exception as e:
            self.logger.error(f"❌ 连接VORTEX失败: {e}")
            self.tm.test_results['steps'].append({
                'step': 1,
                'description': '连接VORTEX Client应用',
                'status': 'FAIL',
                'error': str(e)
            })
            return False

    def test_step_2_locate_target_window(self):
        """步骤2: 定位目标子窗口"""
        self.logger.info("\n--- 步骤2: 定位目标子窗口 ---")
        try:
            vortex_window = self.dlg.window(
                title_re=".*RE小房间-常规.*",
                control_type="Window"
            )
            vortex_window.wait('visible', timeout=5)
            
            self.logger.info(f"✅ 定位到目标窗口: {vortex_window.window_text()}")
            self.vortex_window = vortex_window
            
            self.tm.test_results['steps'].append({
                'step': 2,
                'description': '定位目标子窗口',
                'status': 'PASS',
                'details': f"窗口标题: {vortex_window.window_text()}"
            })
            return True
            
        except ElementNotFoundError as e:
            self.logger.error(f"❌ 目标子窗口定位失败: {e}")
            self.tm.test_results['steps'].append({
                'step': 2,
                'description': '定位目标子窗口',
                'status': 'FAIL',
                'error': str(e)
            })
            return False

    def test_step_3_execute_conversion(self):
        """步骤3: 执行格式转换操作"""
        self.logger.info("\n--- 步骤3: 执行格式转换操作 ---")
        
        steps = [
            ('点击【导出】按钮', 'Button', '导出', None),
            ('点击【点云】选项', 'Pane', '点云', 'option_window'),
            ('选择e57格式', 'RadioButton', 'e57', 'export_window'),
            ('启用点云抽稀', 'CheckBox', '启用', None),
            ('选择合并输出', 'RadioButton', '合并', None),
            ('选择反射率贴图', 'RadioButton', '反射率', None),
            ('启用点云降噪', 'CheckBox', '点云降噪', None),
            ('启用厚度优化', 'CheckBox', '点云厚度优化', None),
            ('点击导出按钮', 'Pane', '导出', None),
        ]
        
        for i, (desc, ctrl_type, title, window_var) in enumerate(steps, 3):
            try:
                if window_var and hasattr(self, window_var):
                    window = getattr(self, window_var)
                else:
                    window = self.export_window if hasattr(self, 'export_window') else self.vortex_window
                
                # 特殊窗口处理
                if desc == '点击【点云】选项':
                    option_window = self.dlg.window(title_re=".*选项.*", control_type="Window")
                    option_window.wait('visible', timeout=5)
                    self.option_window = option_window
                    window = option_window
                    
                elif desc == '选择e57格式':
                    export_window = self.dlg.window(title_re=".*点云导出.*", control_type="Window")
                    export_window.wait('visible', timeout=5)
                    self.export_window = export_window
                    window = export_window
                
                control = window.child_window(control_type=ctrl_type, title=title)
                control.wait('visible enabled', timeout=2)
                
                if ctrl_type == 'CheckBox':
                    control.toggle()
                else:
                    control.click_input()
                
                self.logger.info(f"✅ {desc}成功")
                self.tm.test_results['steps'].append({
                    'step': i,
                    'description': desc,
                    'status': 'PASS'
                })
                
            except Exception as e:
                self.logger.error(f"❌ {desc}失败: {e}")
                self.tm.test_results['steps'].append({
                    'step': i,
                    'description': desc,
                    'status': 'FAIL',
                    'error': str(e)
                })
                return False
        
        return True

    def test_step_4_select_output_path(self):
        """步骤4: 选择输出路径"""
        self.logger.info("\n--- 步骤4: 选择输出路径 ---")
        
        try:
            # 定位浏览文件夹窗口
            browser_window = self.dlg.window(
                title_re=".*浏览文件夹.*",
                control_type="Window"
            )
            browser_window.wait('visible', timeout=5)
            self.logger.info(f"✅ 定位到浏览文件夹窗口")
            
            # 点击此电脑
            browser_window.child_window(control_type="TreeItem", title="此电脑").click_input()
            self.logger.info("✅ 点击【此电脑】成功")
            
            # 点击F盘
            browser_window.child_window(control_type="TreeItem", title="新加卷 (F:)").click_input()
            self.logger.info("✅ 点击【F盘】成功")
            
            # 新建文件夹
            browser_window.child_window(control_type="Button", title="新建文件夹(M)").click_input()
            self.logger.info("✅ 点击【新建文件夹】成功")
            
            # 输入文件夹名
            edit = browser_window.child_window(control_type="Edit", auto_id="1")
            edit.wait('visible enabled', timeout=7)
            folder_name = f"e57+{self.tm.test_config['export_type']}+{self.tm.test_config['texture']}"
            edit.set_text(folder_name)
            self.logger.info(f"✅ 文件夹命名为: {folder_name}")
            
            # 点击确定
            browser_window.child_window(control_type="Button", title="确定").click_input()
            self.logger.info("✅ 点击【确定】成功")
            
            self.tm.test_results['steps'].append({
                'step': 12,
                'description': '选择输出路径并创建文件夹',
                'status': 'PASS',
                'details': f"输出路径: F:\\{folder_name}"
            })
            return True
            
        except Exception as e:
            self.logger.error(f"❌ 选择输出路径失败: {e}")
            self.tm.test_results['steps'].append({
                'step': 12,
                'description': '选择输出路径',
                'status': 'FAIL',
                'error': str(e)
            })
            return False

    def test_step_5_monitor_conversion_process(self):
        """步骤5: 监控转换过程"""
        self.logger.info("\n--- 步骤5: 监控转换过程 ---")
        
        start_time = time.time()
        self.logger.info(f"⏱️ 格式转换计时开始: {datetime.fromtimestamp(start_time).strftime('%Y-%m-%d %H:%M:%S')}")
        
        timeout = self.tm.test_config['timeout']
        check_interval = self.tm.test_config['check_interval']
        convert_success = False
        success_window = None
        
        while not convert_success and (time.time() - start_time) < timeout:
            try:
                success_window = self.dlg.window(auto_id="MessageForm")
                if success_window.exists() and success_window.is_visible():
                    convert_success = True
                    self.logger.info("🎉 检测到【格式转换成功】窗口出现！")
                    
                    # 尝试关闭成功窗口
                    self.close_success_window(success_window)
                    break
                    
                elapsed = time.time() - start_time
                print(f"🔄 等待格式转换完成... 已耗时：{elapsed:.1f}秒", end="\r")
                
            except ElementNotFoundError:
                elapsed = time.time() - start_time
                print(f"🔄 等待格式转换完成... 已耗时：{elapsed:.1f}秒", end="\r")
            
            time.sleep(check_interval)
        
        end_time = time.time()
        duration = end_time - start_time
        
        if convert_success:
            self.logger.info(f"✅ 格式转换完成！总耗时: {duration:.2f}秒")
            self.tm.test_results['steps'].append({
                'step': 13,
                'description': '监控转换过程',
                'status': 'PASS',
                'details': f"转换成功，耗时: {duration:.2f}秒"
            })
            return True, duration
        else:
            self.logger.error(f"❌ 格式转换超时！最大等待{timeout}秒，实际耗时{duration:.2f}秒")
            self.tm.test_results['steps'].append({
                'step': 13,
                'description': '监控转换过程',
                'status': 'FAIL',
                'error': f"转换超时，耗时: {duration:.2f}秒"
            })
            return False, duration

    def close_success_window(self, window):
        """关闭转换成功窗口"""
        self.logger.info("尝试关闭转换成功窗口...")
        
        methods = [
            ("发送回车键", lambda: window.type_keys('{ENTER}')),
            ("发送空格键", lambda: window.type_keys(' ')),
            ("调用close()方法", lambda: window.close()),
            ("发送Alt+F4", lambda: window.type_keys('%{F4}')),
        ]
        
        for method_name, method in methods:
            try:
                method()
                self.logger.info(f"✅ 使用{method_name}关闭窗口成功")
                time.sleep(1)
                if not window.exists():
                    return True
            except Exception as e:
                self.logger.debug(f"⚠️ {method_name}失败: {e}")
        
        return False

    def teardown(self):
        """测试后清理"""
        self.logger.info("\n--- 测试后清理 ---")
        
        # 记录结束时间
        self.tm.test_results['end_time'] = datetime.now().isoformat()
        
        # 计算总耗时
        if self.tm.test_results['start_time'] and self.tm.test_results['end_time']:
            start = datetime.fromisoformat(self.tm.test_results['start_time'])
            end = datetime.fromisoformat(self.tm.test_results['end_time'])
            self.tm.test_results['duration'] = (end - start).total_seconds()
        
        # 判断测试状态
        failed_steps = [step for step in self.tm.test_results['steps'] if step.get('status') == 'FAIL']
        if failed_steps:
            self.tm.test_results['status'] = 'FAIL'
            self.tm.test_results['errors'] = [step.get('error', '未知错误') for step in failed_steps]
        else:
            self.tm.test_results['status'] = 'PASS'

# ==================== 报告生成 ====================
class TestReporter:
    @staticmethod
    def generate_html_report(test_results):
        """生成HTML测试报告"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_dir = "reports"
        if not os.path.exists(report_dir):
            os.makedirs(report_dir)
        
        report_file = os.path.join(report_dir, f"test_report_{timestamp}.html")
        
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>点云格式转换测试报告</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 40px; }}
                .header {{ background-color: #f0f0f0; padding: 20px; border-radius: 5px; }}
                .summary {{ margin: 20px 0; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }}
                .pass {{ color: green; }}
                .fail {{ color: red; }}
                table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
                .step-pass {{ background-color: #d4edda; }}
                .step-fail {{ background-color: #f8d7da; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h1>点云格式转换自动化测试报告</h1>
                <p>生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            </div>
            
            <div class="summary">
                <h2>测试摘要</h2>
                <p><strong>测试名称:</strong> {test_results['test_name']}</p>
                <p><strong>测试状态:</strong> <span class="{test_results['status'].lower()}">{test_results['status']}</span></p>
                <p><strong>开始时间:</strong> {test_results['start_time']}</p>
                <p><strong>结束时间:</strong> {test_results['end_time']}</p>
                <p><strong>总耗时:</strong> {test_results['duration']:.2f}秒</p>
            </div>
            
            <h2>测试步骤详情</h2>
            <table>
                <tr>
                    <th>步骤</th>
                    <th>描述</th>
                    <th>状态</th>
                    <th>详情/错误</th>
                </tr>
        """
        
        for step in test_results['steps']:
            status_class = 'step-pass' if step['status'] == 'PASS' else 'step-fail'
            html_content += f"""
                <tr class="{status_class}">
                    <td>{step['step']}</td>
                    <td>{step['description']}</td>
                    <td>{step['status']}</td>
                    <td>{step.get('details', step.get('error', ''))}</td>
                </tr>
            """
        
        html_content += """
            </table>
        </body>
        </html>
        """
        
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return report_file

    @staticmethod
    def generate_json_report(test_results):
        """生成JSON测试报告"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_dir = "reports"
        if not os.path.exists(report_dir):
            os.makedirs(report_dir)
        
        report_file = os.path.join(report_dir, f"test_report_{timestamp}.json")
        
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(test_results, f, indent=2, ensure_ascii=False)
        
        return report_file

# ==================== 主执行流程 ====================
def main():
    # 初始化测试管理器
    test_manager = PointCloudConversionTest()
    
    # 初始化测试用例
    test_case = PointCloudConversionTestCase(test_manager)
    
    try:
        # 1. 测试设置
        test_case.setup()
        
        # 2. 执行测试步骤
        steps = [
            test_case.test_step_1_connect_to_vortex,
            test_case.test_step_2_locate_target_window,
            test_case.test_step_3_execute_conversion,
            test_case.test_step_4_select_output_path,
            test_case.test_step_5_monitor_conversion_process
        ]
        
        for step_func in steps:
            if not step_func():
                test_manager.logger.error("测试步骤失败，终止测试")
                break
        
        # 3. 测试清理
        test_case.teardown()
        
        # 4. 生成测试报告
        reporter = TestReporter()
        html_report = reporter.generate_html_report(test_manager.test_results)
        json_report = reporter.generate_json_report(test_manager.test_results)
        
        test_manager.logger.info(f"📊 HTML测试报告已生成: {html_report}")
        test_manager.logger.info(f"📊 JSON测试报告已生成: {json_report}")
        
        # 5. 控制台总结
        print("\n" + "="*60)
        print("测试执行完成！")
        print(f"测试状态: {test_manager.test_results['status']}")
        print(f"总耗时: {test_manager.test_results['duration']:.2f}秒")
        print(f"总步骤: {len(test_manager.test_results['steps'])}")
        print(f"通过步骤: {len([s for s in test_manager.test_results['steps'] if s['status'] == 'PASS'])}")
        print(f"失败步骤: {len([s for s in test_manager.test_results['steps'] if s['status'] == 'FAIL'])}")
        
        if test_manager.test_results['errors']:
            print("\n错误列表:")
            for error in test_manager.test_results['errors']:
                print(f"  - {error}")
        
    except Exception as e:
        test_manager.logger.error(f"测试执行过程中发生未预期的错误: {e}")
        test_manager.test_results['status'] = 'ERROR'
        test_manager.test_results['errors'].append(str(e))
        
        # 生成错误报告
        reporter = TestReporter()
        reporter.generate_html_report(test_manager.test_results)
        reporter.generate_json_report(test_manager.test_results)

if __name__ == "__main__":
    main()