from pywinauto import Application, ElementNotFoundError
import win32gui 
import time
import csv
import os
import json
from datetime import datetime
import logging
import pandas as pd
from typing import Dict, List, Any, Optional
import traceback

# ==================== 配置部分 ====================
class DataDrivenPointCloudTest:
    def __init__(self, config_file: str = "test_config.json"):
        """
        初始化数据驱动测试
        
        Args:
            config_file: 配置文件路径
        """
        self.setup_logging()
        self.load_config(config_file)
        self.setup_directories()
        
        # 存储所有测试用例结果
        self.all_results = {
            "total_cases": 0,
            "passed_cases": 0,
            "failed_cases": 0,
            "error_cases": 0,
            "test_cases": [],
            "start_time": None,
            "end_time": None,
            "total_duration": 0
        }

    def setup_logging(self):
        """配置日志系统"""
        log_dir = "logs"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_dir, f"data_driven_test_{timestamp}.log")
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

    def load_config(self, config_file: str):
        """加载配置文件"""
        default_config = {
            "vortex_window_title": "VORTEX Client",
            "target_window_pattern": ".*建模_20251231025100.*",
            "timeout": 1200,
            "check_interval": 0.5,
             "csv_file": r"D:\study\test_vortexclient\test_cases\all_test_cases_complete.csv",
            "output_base_dir": "F:\\",
            "backend": "uia",
            "wait_after_enable_thinning": 1.0,
        }
        
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                self.logger.info(f"已加载配置文件: {config_file}")
            else:
                self.config = default_config
                self.logger.info("使用默认配置")
        except Exception as e:
            self.logger.error(f"加载配置文件失败: {e}")
            self.config = default_config

    def setup_directories(self):
        """创建必要的目录"""
        directories = ["logs", "reports", "screenshots", "outputs", "test_cases"]
        for directory in directories:
            if not os.path.exists(directory):
                os.makedirs(directory)

# ==================== CSV数据读取器 ====================
class CSVDataReader:
    @staticmethod
    def read_test_cases(csv_file: str) -> List[Dict[str, Any]]:
        """
        从CSV文件读取测试用例
        
        Args:
            csv_file: CSV文件路径
            
        Returns:
            测试用例列表，每个用例是一个字典
        """
        test_cases = []
        
        try:
            df = pd.read_csv(csv_file, encoding='utf-8')
            
            # 确保所有必需的列都存在
            required_columns = ['输出格式', '点云抽稀', '输出类型', '贴图选择', 
                              '点云降噪', '点云厚度优化']
            
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"CSV文件中缺少必需的列: {col}")
            
            # 转换为字典列表
            for index, row in df.iterrows():
                test_case = {
                    "用例ID": row.get('用例ID', f"TC{index+1:03d}"),
                    "输出格式": str(row['输出格式']).strip(),
                    "点云抽稀": str(row['点云抽稀']).strip(),
                    "体素抽稀": str(row.get('体素抽稀', '')).strip() if pd.notna(row.get('体素抽稀')) else "",
                    "随机抽稀": str(row.get('随机抽稀', '')).strip() if pd.notna(row.get('随机抽稀')) else "",
                    "输出类型": str(row['输出类型']).strip(),
                    "贴图选择": str(row['贴图选择']).strip(),
                    "点云降噪": str(row['点云降噪']).strip(),
                    "点云厚度优化": str(row['点云厚度优化']).strip(),
                    "预期结果": row.get('预期结果', '成功'),
                    "备注": row.get('备注', ''),
                    "row_index": index + 2  # 包括标题行
                }
                
                # 验证抽稀方式逻辑
                if test_case["点云抽稀"] == "启用":
                    if not test_case["体素抽稀"] and not test_case["随机抽稀"]:
                        raise ValueError(f"用例 {test_case['用例ID']} 启用了点云抽稀，但未指定抽稀方式")
                    if test_case["体素抽稀"] not in ["启用", "不启用", ""]:
                        raise ValueError(f"用例 {test_case['用例ID']} 体素抽稀值无效: {test_case['体素抽稀']}")
                    if test_case["随机抽稀"] not in ["启用", "不启用", ""]:
                        raise ValueError(f"用例 {test_case['用例ID']} 随机抽稀值无效: {test_case['随机抽稀']}")
                
                if test_case["点云抽稀"] != "启用" and (test_case["体素抽稀"] or test_case["随机抽稀"]):
                    raise ValueError(f"用例 {test_case['用例ID']} 未启用点云抽稀，但指定了抽稀方式")
                
                test_cases.append(test_case)
            
            return test_cases
            
        except Exception as e:
            raise Exception(f"读取CSV文件失败: {e}")

# ==================== 测试用例执行器 ====================
class TestCaseExecutor:
    def __init__(self, test_manager, test_case: Dict[str, Any]):
        self.tm = test_manager
        self.logger = test_manager.logger
        self.test_case = test_case
        self.app = None
        self.dlg = None
        self.vortex_window = None
        
        # 计时相关
        self.conversion_start_time = None
        self.conversion_end_time = None
        self.conversion_duration = None
        
        # 测试结果
        self.result = {
            "用例ID": test_case["用例ID"],
            "配置": test_case,
            "状态": "未执行",
            "步骤": [],
            "开始时间": None,
            "结束时间": None,
            "持续时间": None,
            "转换开始时间": None,
            "转换结束时间": None,
            "转换耗时": None,
            "错误信息": None,
            "输出文件夹": None
        }

    def execute(self) -> bool:
        """执行单个测试用例"""
        self.result["开始时间"] = datetime.now().isoformat()
        
        try:
            self.logger.info(f"\n{'='*60}")
            self.logger.info(f"开始执行测试用例: {self.test_case['用例ID']}")
            self.logger.info(f"配置: {json.dumps(self.test_case, indent=2, ensure_ascii=False)}")
            
            # 执行所有步骤
            steps = [
                self._connect_to_vortex,
                self._locate_target_window,
                self._click_export_button,
                self._select_point_cloud_option,
                self._configure_export_settings,
                self._select_output_path,
                self._monitor_conversion_process
            ]
            
            for step_func in steps:
                if not step_func():
                    self.result["状态"] = "失败"
                    break
            else:
                self.result["状态"] = "通过"
            
        except Exception as e:
            self.logger.error(f"测试用例执行异常: {e}")
            self.logger.error(traceback.format_exc())
            self.result["状态"] = "错误"
            self.result["错误信息"] = str(e)
        
        finally:
            self.result["结束时间"] = datetime.now().isoformat()
            if self.result["开始时间"] and self.result["结束时间"]:
                start = datetime.fromisoformat(self.result["开始时间"])
                end = datetime.fromisoformat(self.result["结束时间"])
                self.result["持续时间"] = (end - start).total_seconds()
            
            # 记录转换耗时
            if self.conversion_start_time and self.conversion_end_time:
                self.result["转换开始时间"] = self.conversion_start_time.isoformat()
                self.result["转换结束时间"] = self.conversion_end_time.isoformat()
                self.result["转换耗时"] = (self.conversion_end_time - self.conversion_start_time).total_seconds()
            
            self.logger.info(f"测试用例 {self.test_case['用例ID']} 执行完成，状态: {self.result['状态']}")
            
            # 清理资源
            self._cleanup()
            
            return self.result["状态"] == "通过"

    def _connect_to_vortex(self) -> bool:
        """连接到VORTEX应用"""
        try:
            hwnd = win32gui.FindWindow(None, self.tm.config["vortex_window_title"])
            if not hwnd:
                raise Exception(f"未找到标题为'{self.tm.config['vortex_window_title']}'的窗口")
            
            self.app = Application(backend=self.tm.config["backend"]).connect(handle=hwnd)
            self.dlg = self.app.window(handle=hwnd)
            self.dlg.wait('visible enabled', timeout=10)
            self.dlg.set_focus()
            
            self._add_step("连接VORTEX", "通过", f"句柄: {hex(hwnd)}")
            return True
            
        except Exception as e:
            self._add_step("连接VORTEX", "失败", str(e))
            return False

    def _locate_target_window(self) -> bool:
        """定位目标窗口"""
        try:
            self.vortex_window = self.dlg.window(
                title_re=self.tm.config["target_window_pattern"],
                control_type="Window"
            )
            self.vortex_window.wait('visible', timeout=5)
            
            self._add_step("定位目标窗口", "通过", 
                          f"窗口标题: {self.vortex_window.window_text()}")
            return True
            
        except ElementNotFoundError as e:
            self._add_step("定位目标窗口", "失败", str(e))
            return False

    def _click_export_button(self) -> bool:
        """点击导出按钮"""
        try:
            export_button = self.vortex_window.child_window(
                control_type="Button",
                title="导出"
            )
            export_button.wait('visible enabled', timeout=5)
            export_button.click_input()
            
            self._add_step("点击导出按钮", "通过")
            return True
            
        except Exception as e:
            self._add_step("点击导出按钮", "失败", str(e))
            return False

    def _select_point_cloud_option(self) -> bool:
        """选择点云选项"""
        try:
            # 定位选项窗口
            option_window = self.dlg.window(
                title_re=".*选项.*",
                control_type="Window"
            )
            option_window.wait('visible', timeout=5)
            
            # 点击点云选项
            point_cloud_option = option_window.child_window(
                control_type="Pane",
                title="点云"
            )
            point_cloud_option.wait('visible enabled', timeout=2)
            point_cloud_option.click_input()
            
            self._add_step("选择点云选项", "通过")
            return True
            
        except Exception as e:
            self._add_step("选择点云选项", "失败", str(e))
            return False

    def _configure_export_settings(self) -> bool:
        """配置导出设置（根据CSV参数）"""
        try:
            # 定位导出窗口
            export_window = self.dlg.window(
                title_re=".*点云导出.*",
                control_type="Window"
            )
            export_window.wait('visible', timeout=5)
            
            # 1. 选择输出格式
            format_mapping = {
                "pts": "pts",
                "e57": "e57",
                "las": "las"
            }
            format_title = format_mapping.get(self.test_case["输出格式"], "e57")
            self._select_radio_button(export_window, format_title, "输出格式")
            
            # 2. 配置点云抽稀
            if self.test_case["点云抽稀"] == "启用":
                # 点击启用复选框
                self._toggle_checkbox(export_window, "启用", "点云抽稀")
                
                # 等待抽稀选项出现
                time.sleep(self.tm.config.get("wait_after_enable_thinning", 1.0))
                
                # 配置体素抽稀 - 如果是启用状态，就点击单选按钮
                if self.test_case["体素抽稀"] == "启用":
                    try:
                        voxel_radio = export_window.child_window(
                            control_type="RadioButton",
                            title="体素抽稀"
                        )
                        voxel_radio.wait('visible enabled', timeout=2)
                        voxel_radio.click_input()
                        self._add_step("配置体素抽稀", "通过", "启用体素抽稀")
                    except Exception as e:
                        self.logger.error(f"点击体素抽稀失败: {e}")
                        self._add_step("配置体素抽稀", "失败", str(e))
                        return False
                
                # 配置随机抽稀 - 如果是启用状态，就点击单选按钮
                if self.test_case["随机抽稀"] == "启用":
                    try:
                        random_radio = export_window.child_window(
                            control_type="RadioButton",
                            title="随机抽稀"
                        )
                        random_radio.wait('visible enabled', timeout=2)
                        random_radio.click_input()
                        self._add_step("配置随机抽稀", "通过", "启用随机抽稀")
                    except Exception as e:
                        self.logger.error(f"点击随机抽稀失败: {e}")
                        self._add_step("配置随机抽稀", "失败", str(e))
                        return False
            
            # 3. 选择输出类型
            output_type_mapping = {
                "单站": "单站",
                "合并": "合并",
                "单站+合并": "单站+合并"
            }
            output_type_title = output_type_mapping.get(self.test_case["输出类型"], "合并")
            self._select_radio_button(export_window, output_type_title, "输出类型")
            
            # 4. 选择贴图
            texture_mapping = {
                "灰阶图": "灰阶图",
                "反射率": "反射率",
                "反射率+彩图": "反射率+彩图",
                "反射率+灰阶图": "反射率+灰阶图"
            }
            texture_title = texture_mapping.get(self.test_case["贴图选择"], "反射率")
            self._select_radio_button(export_window, texture_title, "贴图选择")
            
            # 5. 配置点云降噪
            if self.test_case["点云降噪"] == "启用":
                self._toggle_checkbox(export_window, "点云降噪", "点云降噪")
            
            # 6. 配置点云厚度优化
            if self.test_case["点云厚度优化"] == "启用":
                self._toggle_checkbox(export_window, "点云厚度优化", "点云厚度优化")
            
            # 7. 点击导出按钮
            export_button = export_window.child_window(
                control_type="Pane",
                title="导出",
                auto_id="uiButton3"
            )
            export_button.wait('visible enabled', timeout=2)
            export_button.click_input()
            
            self._add_step("配置导出设置", "通过")
            return True
            
        except Exception as e:
            self._add_step("配置导出设置", "失败", str(e))
            return False

    def _select_radio_button(self, parent_window, title: str, step_name: str):
        """选择单选按钮"""
        radio_button = parent_window.child_window(
            control_type="RadioButton",
            title=title
        )
        radio_button.wait('visible enabled', timeout=2)
        radio_button.click_input()
        self._add_step(f"选择{step_name}", "通过", f"选择: {title}")

    def _toggle_checkbox(self, parent_window, title: str, step_name: str):
        """切换复选框状态"""
        checkbox = parent_window.child_window(
            control_type="CheckBox",
            title=title
        )
        checkbox.wait('visible enabled', timeout=2)
        
        # 先检查当前状态
        try:
            # 尝试获取选中状态
            if hasattr(checkbox, 'get_toggle_state'):
                current_state = checkbox.get_toggle_state()
                # 如果已经是选中状态，不需要切换
                if current_state == 1:  # 1表示选中
                    self._add_step(f"配置{step_name}", "通过", f"已启用: {title}")
                    return
        except:
            pass  # 如果获取状态失败，继续执行toggle
        
        checkbox.toggle()
        self._add_step(f"配置{step_name}", "通过", f"状态: {title}")

    def _select_output_path(self) -> bool:
        """选择输出路径"""
        try:
            # 生成有意义的文件夹名称
            folder_name = self._generate_folder_name()
            self.result["输出文件夹"] = folder_name
            
            # 定位浏览文件夹窗口
            browser_window = self.dlg.window(
                title_re=".*浏览文件夹.*",
                control_type="Window"
            )
            browser_window.wait('visible', timeout=5)
            
            # 点击此电脑
            browser_window.child_window(control_type="TreeItem", title="此电脑").click_input()
            time.sleep(1)
            
            # 点击D盘
            browser_window.child_window(control_type="TreeItem", title="Data (D:)").click_input()
            time.sleep(1)
            
            # 新建文件夹
            browser_window.child_window(control_type="Button", title="新建文件夹(M)").click_input()
            time.sleep(2)
            
            # 输入文件夹名
            edit = browser_window.child_window(control_type="Edit", auto_id="1")
            edit.wait('visible enabled', timeout=7)
            edit.set_text(folder_name)
            
            # 点击确定 - 开始记录转换时间
            ok_button = browser_window.child_window(control_type="Button", title="确定")
            ok_button.wait('visible enabled', timeout=2)
            ok_button.click_input()
            
            # 记录转换开始时间
            self.conversion_start_time = datetime.now()
            self.result["转换开始时间"] = self.conversion_start_time.isoformat()
            self.logger.info(f"⏱️ 转换计时开始: {self.conversion_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
            
            self._add_step("选择输出路径", "通过", f"文件夹: {folder_name}")
            return True
            
        except Exception as e:
            self._add_step("选择输出路径", "失败", str(e))
            return False

    def _generate_folder_name(self) -> str:
        """生成有意义的文件夹名称"""
        parts = []
        
        # 输出格式
        parts.append(f"格式-{self.test_case['输出格式']}")
        
        # 点云抽稀
        if self.test_case["点云抽稀"] == "启用":
            thinning_parts = []
            if self.test_case["体素抽稀"] == "启用":
                thinning_parts.append("体素")
            if self.test_case["随机抽稀"] == "启用":
                thinning_parts.append("随机")
            if thinning_parts:
                parts.append(f"抽稀-{'+'.join(thinning_parts)}")
            else:
                parts.append("抽稀-无")
        else:
            parts.append("抽稀-否")
        
        # 输出类型
        parts.append(f"输出-{self.test_case['输出类型']}")
        
        # 贴图选择
        texture_map = {
            "灰阶图": "灰度",
            "反射率": "反射",
            "反射率+彩图": "反射+彩图",
            "反射率+灰阶图": "反射+灰度"
        }
        texture_short = texture_map.get(self.test_case["贴图选择"], "反射")
        parts.append(f"贴图-{texture_short}")
        
        # 点云降噪
        parts.append(f"降噪-{self.test_case['点云降噪']}")
        
        # 点云厚度优化
        parts.append(f"厚度-{self.test_case['点云厚度优化']}")
        
        # 时间戳（避免重名）
        timestamp = datetime.now().strftime("%m%d%H%M")
        parts.append(timestamp)
        
        # 组合所有部分
        folder_name = "_".join(parts)
        
        # 限制长度（Windows路径最大260字符）
        max_length = 50
        if len(folder_name) > max_length:
            folder_name = folder_name[:max_length]
        
        return folder_name

    def _monitor_conversion_process(self) -> bool:
        """监控转换过程"""
        try:
            timeout = self.tm.config["timeout"]
            check_interval = self.tm.config["check_interval"]
            convert_success = False
            
            self.logger.info(f"开始监控转换过程，超时时间: {timeout}秒")
            
            while not convert_success and ((datetime.now() - self.conversion_start_time).total_seconds() if self.conversion_start_time else 0) < timeout:
                try:
                    success_window = self.dlg.window(auto_id="MessageForm")
                    if success_window.exists() and success_window.is_visible():
                        convert_success = True
                        
                        # 记录转换结束时间
                        self.conversion_end_time = datetime.now()
                        conversion_duration = (self.conversion_end_time - self.conversion_start_time).total_seconds()
                        
                        # 获取窗口信息
                        try:
                            window_text = success_window.window_text()
                            self.logger.info(f"检测到成功窗口: {window_text}")
                        except:
                            pass
                        
                        self.logger.info(f"✅ 格式转换成功！转换耗时: {conversion_duration:.2f}秒")
                        
                        # 尝试关闭成功窗口
                        self._close_success_window(success_window)
                        break
                    
                    if self.conversion_start_time:
                        elapsed = (datetime.now() - self.conversion_start_time).total_seconds()
                        print(f"等待转换完成... 已耗时: {elapsed:.1f}秒", end="\r")
                    
                except ElementNotFoundError:
                    if self.conversion_start_time:
                        elapsed = (datetime.now() - self.conversion_start_time).total_seconds()
                        print(f"等待转换完成... 已耗时: {elapsed:.1f}秒", end="\r")
                
                time.sleep(check_interval)
            
            if not convert_success:
                elapsed_time = (datetime.now() - self.conversion_start_time).total_seconds() if self.conversion_start_time else 0
                self._add_step("监控转换过程", "失败", f"超时，耗时: {elapsed_time:.2f}秒")
                return False
            else:
                conversion_duration = (self.conversion_end_time - self.conversion_start_time).total_seconds()
                self._add_step("监控转换过程", "通过", f"转换耗时: {conversion_duration:.2f}秒")
                return True
                
        except Exception as e:
            self._add_step("监控转换过程", "失败", str(e))
            return False

    def _close_success_window(self, window) -> bool:
        """关闭成功窗口"""
        try:
            # 尝试查找并点击确定按钮
            try:
                # 先尝试通过标题查找
                ok_button = window.child_window(
                    control_type="Button",
                    title="确定"
                )
                if ok_button.exists():
                    ok_button.wait('visible enabled', timeout=2)
                    ok_button.click_input()
                    self.logger.info("点击确定按钮成功")
                    return True
            except:
                pass
            
            # 尝试多种关闭方式
            methods = [
                ("回车键", lambda: window.type_keys('{ENTER}')),
                ("空格键", lambda: window.type_keys(' ')),
                ("close方法", lambda: window.close()),
                ("Alt+F4", lambda: window.type_keys('%{F4}')),
            ]
            
            for method_name, method in methods:
                try:
                    method()
                    time.sleep(1)
                    if not window.exists():
                        self.logger.info(f"成功关闭窗口: {method_name}")
                        return True
                except:
                    continue
            
            # 如果以上方法都失败，尝试点击窗口任意位置
            try:
                window.click_input()
                self.logger.info("点击窗口任意位置")
                return True
            except:
                pass
            
            return False
            
        except Exception as e:
            self.logger.warning(f"关闭窗口失败: {e}")
            return False

    def _add_step(self, description: str, status: str, details: str = ""):
        """添加测试步骤"""
        step = {
            "步骤": description,
            "状态": status,
            "详情": details,
            "时间": datetime.now().isoformat()
        }
        self.result["步骤"].append(step)

    def _cleanup(self):
        """清理资源"""
        try:
            # 如果应用还在运行，尝试关闭所有子窗口
            if hasattr(self, 'app') and self.app:
                # 这里可以添加更多清理逻辑
                pass
        except:
            pass

# ==================== 报告生成器 ====================
class DataDrivenTestReporter:
    @staticmethod
    def generate_html_report(all_results: Dict[str, Any], output_file: str = None):
        """生成HTML测试报告"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if not output_file:
            output_file = f"reports/data_driven_report_{timestamp}.html"
        
        # 计算统计信息
        total = all_results["total_cases"]
        passed = all_results["passed_cases"]
        failed = all_results["failed_cases"]
        error = all_results["error_cases"]
        pass_rate = (passed / total * 100) if total > 0 else 0
        
        # 计算平均转换时间
        conversion_times = []
        for test_case in all_results["test_cases"]:
            if test_case.get("转换耗时"):
                conversion_times.append(test_case["转换耗时"])
        
        avg_conversion_time = sum(conversion_times) / len(conversion_times) if conversion_times else 0
        
        # 生成状态颜色
        status_colors = {
            "通过": "#28a745",  # 绿色
            "失败": "#dc3545",  # 红色
            "错误": "#ffc107",  # 黄色
            "未执行": "#6c757d"  # 灰色
        }
        
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>点云格式转换数据驱动测试报告</title>
            <style>
                body {{ font-family: 'Microsoft YaHei', Arial, sans-serif; margin: 40px; }}
                .header {{ background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px; }}
                .summary {{ margin: 30px 0; padding: 25px; background: #f8f9fa; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
                .stats {{ display: flex; justify-content: space-around; margin: 20px 0; }}
                .stat-card {{ background: white; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
                .stat-pass {{ border-left: 5px solid #28a745; }}
                .stat-fail {{ border-left: 5px solid #dc3545; }}
                .stat-error {{ border-left: 5px solid #ffc107; }}
                .stat-total {{ border-left: 5px solid #007bff; }}
                .stat-time {{ border-left: 5px solid #17a2b8; }}
                table {{ width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 10px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
                th, td {{ border: 1px solid #dee2e6; padding: 12px; text-align: left; }}
                th {{ background-color: #f2f2f2; font-weight: bold; }}
                .test-case-row:hover {{ background-color: #f8f9fa; }}
                .progress-bar {{ height: 20px; background: #e9ecef; border-radius: 10px; overflow: hidden; margin: 10px 0; }}
                .progress-fill {{ height: 100%; background: #28a745; }}
                .toggle-details {{ cursor: pointer; color: #007bff; }}
                .test-details {{ display: none; padding: 15px; background: #f8f9fa; border-radius: 5px; margin: 10px 0; }}
                .config-cell {{ max-width: 80px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
                .config-cell:hover {{ overflow: visible; white-space: normal; background: white; z-index: 100; position: relative; }}
                .time-cell {{ font-family: monospace; }}
            </style>
            <script>
                function toggleDetails(caseId) {{
                    var details = document.getElementById('details-' + caseId);
                    var button = document.getElementById('button-' + caseId);
                    if (details.style.display === 'none') {{
                        details.style.display = 'block';
                        button.textContent = '收起详情';
                    }} else {{
                        details.style.display = 'none';
                        button.textContent = '查看详情';
                    }}
                }}
            </script>
        </head>
        <body>
            <div class="header">
                <h1>📊 点云格式转换数据驱动测试报告</h1>
                <p>生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            </div>
            
            <div class="summary">
                <h2>📈 测试摘要</h2>
                <div class="stats">
                    <div class="stat-card stat-total">
                        <h3>总计</h3>
                        <h2>{total}</h2>
                    </div>
                    <div class="stat-card stat-pass">
                        <h3>通过</h3>
                        <h2>{passed}</h2>
                    </div>
                    <div class="stat-card stat-fail">
                        <h3>失败</h3>
                        <h2>{failed}</h2>
                    </div>
                    <div class="stat-card stat-error">
                        <h3>错误</h3>
                        <h2>{error}</h2>
                    </div>
                    <div class="stat-card stat-time">
                        <h3>平均转换时间</h3>
                        <h2>{avg_conversion_time:.1f}秒</h2>
                    </div>
                </div>
                
                <div>
                    <h4>通过率</h4>
                    <div class="progress-bar">
                        <div class="progress-fill" style="width: {pass_rate}%"></div>
                    </div>
                    <p>{pass_rate:.2f}% ({passed}/{total})</p>
                </div>
                
                <p><strong>总耗时:</strong> {all_results['total_duration']:.2f}秒</p>
                <p><strong>开始时间:</strong> {all_results['start_time']}</p>
                <p><strong>结束时间:</strong> {all_results['end_time']}</p>
            </div>
            
            <h2>📋 测试用例详情</h2>
            <table>
                <tr>
                    <th>用例ID</th>
                    <th>输出格式</th>
                    <th>点云抽稀</th>
                    <th>体素抽稀</th>
                    <th>随机抽稀</th>
                    <th>输出类型</th>
                    <th>贴图选择</th>
                    <th>点云降噪</th>
                    <th>厚度优化</th>
                    <th>状态</th>
                    <th>转换耗时(秒)</th>
                    <th>操作</th>
                </tr>
        """
        
        for test_case in all_results["test_cases"]:
            # 获取状态对应的颜色
            status_color = status_colors.get(test_case['状态'], "#6c757d")
            details_id = test_case["用例ID"].replace(" ", "_").replace(".", "_")
            
            # 格式化转换时间
            conversion_time = f"{test_case['转换耗时']:.1f}" if test_case.get("转换耗时") else "N/A"
            
            html_content += f"""
                <tr class="test-case-row">
                    <td>{test_case['用例ID']}</td>
                    <td class="config-cell">{test_case['配置']['输出格式']}</td>
                    <td class="config-cell">{test_case['配置']['点云抽稀']}</td>
                    <td class="config-cell">{test_case['配置'].get('体素抽稀', 'N/A')}</td>
                    <td class="config-cell">{test_case['配置'].get('随机抽稀', 'N/A')}</td>
                    <td class="config-cell">{test_case['配置']['输出类型']}</td>
                    <td class="config-cell" title="{test_case['配置']['贴图选择']}">{test_case['配置']['贴图选择'][:10]}{'...' if len(test_case['配置']['贴图选择']) > 10 else ''}</td>
                    <td class="config-cell">{test_case['配置']['点云降噪']}</td>
                    <td class="config-cell">{test_case['配置']['点云厚度优化']}</td>
                    <td style="color: {status_color}; font-weight: bold;">{test_case['状态']}</td>
                    <td class="time-cell">{conversion_time}</td>
                    <td><button id="button-{details_id}" class="toggle-details" onclick="toggleDetails('{details_id}')">查看详情</button></td>
                </tr>
                <tr>
                    <td colspan="12">
                        <div class="test-details" id="details-{details_id}">
                            <h4>测试配置详情:</h4>
                            <ul>
                                <li><strong>用例ID:</strong> {test_case['用例ID']}</li>
                                <li><strong>输出格式:</strong> {test_case['配置']['输出格式']}</li>
                                <li><strong>点云抽稀:</strong> {test_case['配置']['点云抽稀']}</li>
                                <li><strong>体素抽稀:</strong> {test_case['配置'].get('体素抽稀', 'N/A')}</li>
                                <li><strong>随机抽稀:</strong> {test_case['配置'].get('随机抽稀', 'N/A')}</li>
                                <li><strong>输出类型:</strong> {test_case['配置']['输出类型']}</li>
                                <li><strong>贴图选择:</strong> {test_case['配置']['贴图选择']}</li>
                                <li><strong>点云降噪:</strong> {test_case['配置']['点云降噪']}</li>
                                <li><strong>点云厚度优化:</strong> {test_case['配置']['点云厚度优化']}</li>
                                <li><strong>预期结果:</strong> {test_case['配置'].get('预期结果', 'N/A')}</li>
                                <li><strong>备注:</strong> {test_case['配置'].get('备注', 'N/A')}</li>
                                <li><strong>输出文件夹:</strong> {test_case.get('输出文件夹', 'N/A')}</li>
                                <li><strong>转换开始时间:</strong> {test_case.get('转换开始时间', 'N/A')}</li>
                                <li><strong>转换结束时间:</strong> {test_case.get('转换结束时间', 'N/A')}</li>
                                <li><strong>转换耗时:</strong> {conversion_time}秒</li>
                            </ul>
                            
                            <h4>测试步骤:</h4>
                            <table>
                                <tr>
                                    <th>步骤</th>
                                    <th>状态</th>
                                    <th>详情</th>
                                    <th>时间</th>
                                </tr>
            """
            
            for step in test_case.get("步骤", []):
                step_color = status_colors.get(step['状态'], "#6c757d")
                html_content += f"""
                                <tr>
                                    <td>{step['步骤']}</td>
                                    <td style="color: {step_color}; font-weight: bold;">{step['状态']}</td>
                                    <td>{step.get('详情', '')}</td>
                                    <td>{step['时间']}</td>
                                </tr>
                """
            
            if test_case.get("错误信息"):
                html_content += f"""
                                <tr>
                                    <td colspan="4" style="color: #dc3545;">
                                        <strong>错误信息:</strong> {test_case['错误信息']}
                                    </td>
                                </tr>
                """
            
            html_content += """
                            </table>
                        </div>
                    </td>
                </tr>
            """
        
        html_content += """
            </table>
        </body>
        </html>
        """
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return output_file

    @staticmethod
    def generate_csv_summary(all_results: Dict[str, Any]):
        """生成CSV汇总报告"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"reports/test_summary_{timestamp}.csv"
        
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # 写入头部
            writer.writerow(['用例ID', '输出格式', '点云抽稀', '体素抽稀', '随机抽稀', 
                           '输出类型', '贴图选择', '点云降噪', '点云厚度优化', 
                           '状态', '转换耗时(秒)', '转换开始时间', '转换结束时间',
                           '总耗时(秒)', '开始时间', '结束时间', '输出文件夹', '备注'])
            
            # 写入数据
            for test_case in all_results["test_cases"]:
                writer.writerow([
                    test_case["用例ID"],
                    test_case["配置"]["输出格式"],
                    test_case["配置"]["点云抽稀"],
                    test_case["配置"].get("体素抽稀", ""),
                    test_case["配置"].get("随机抽稀", ""),
                    test_case["配置"]["输出类型"],
                    test_case["配置"]["贴图选择"],
                    test_case["配置"]["点云降噪"],
                    test_case["配置"]["点云厚度优化"],
                    test_case["状态"],
                    f"{test_case['转换耗时']:.2f}" if test_case.get("转换耗时") else "N/A",
                    test_case.get("转换开始时间", ""),
                    test_case.get("转换结束时间", ""),
                    f"{test_case['持续时间']:.2f}" if test_case["持续时间"] else "N/A",
                    test_case["开始时间"],
                    test_case["结束时间"],
                    test_case.get("输出文件夹", ""),
                    test_case["配置"].get("备注", "")
                ])
        
        return output_file

# ==================== 主执行流程 ====================
def main():
    """主执行函数"""
    
    # 1. 初始化测试管理器
    test_manager = DataDrivenPointCloudTest()
    
    # 2. 检查是否有测试用例文件，如果没有则生成
    csv_file = test_manager.config["csv_file"]
    if not os.path.exists(csv_file):
        print(f"⚠️ 未找到测试用例文件: {csv_file}")
        print("正在生成示例测试用例...")
        generator = TestCaseGenerator()
        sample_cases, generated_file = generator.generate_sample_test_cases()
        test_manager.config["csv_file"] = generated_file
        csv_file = generated_file
    
    # 3. 读取CSV测试用例
    try:
        test_cases = CSVDataReader.read_test_cases(csv_file)
        test_manager.logger.info(f"从 {csv_file} 读取到 {len(test_cases)} 个测试用例")
    except Exception as e:
        test_manager.logger.error(f"读取测试用例失败: {e}")
        return
    
    # 4. 初始化结果
    test_manager.all_results["total_cases"] = len(test_cases)
    test_manager.all_results["start_time"] = datetime.now().isoformat()
    
    # 5. 执行所有测试用例
    for i, test_case in enumerate(test_cases):
        test_manager.logger.info(f"\n{'='*60}")
        test_manager.logger.info(f"执行测试用例 {i+1}/{len(test_cases)}: {test_case['用例ID']}")
        
        # 创建执行器
        executor = TestCaseExecutor(test_manager, test_case)
        
        # 执行测试用例
        success = executor.execute()
        
        # 记录结果
        test_manager.all_results["test_cases"].append(executor.result)
        
        if success:
            test_manager.all_results["passed_cases"] += 1
        elif executor.result["状态"] == "失败":
            test_manager.all_results["failed_cases"] += 1
        else:
            test_manager.all_results["error_cases"] += 1
        
        # 短暂暂停，避免过快执行
        if i < len(test_cases) - 1:  # 如果不是最后一个用例
            time.sleep(3)  # 等待3秒，确保前一个用例完全结束
    
    # 6. 完成统计
    test_manager.all_results["end_time"] = datetime.now().isoformat()
    
    if test_manager.all_results["start_time"] and test_manager.all_results["end_time"]:
        start = datetime.fromisoformat(test_manager.all_results["start_time"])
        end = datetime.fromisoformat(test_manager.all_results["end_time"])
        test_manager.all_results["total_duration"] = (end - start).total_seconds()
    
    # 7. 生成报告
    reporter = DataDrivenTestReporter()
    
    html_report = reporter.generate_html_report(test_manager.all_results)
    csv_summary = reporter.generate_csv_summary(test_manager.all_results)
    
    test_manager.logger.info(f"\n{'='*60}")
    test_manager.logger.info("🎉 所有测试用例执行完成！")
    test_manager.logger.info(f"📊 HTML报告: {html_report}")
    test_manager.logger.info(f"📊 CSV汇总: {csv_summary}")
    
    # 8. 控制台总结
    print("\n" + "="*60)
    print("📈 测试执行完成！")
    print(f"总计: {test_manager.all_results['total_cases']}")
    print(f"通过: {test_manager.all_results['passed_cases']}")
    print(f"失败: {test_manager.all_results['failed_cases']}")
    print(f"错误: {test_manager.all_results['error_cases']}")
    
    pass_rate = (test_manager.all_results['passed_cases'] / 
                 test_manager.all_results['total_cases'] * 100) if test_manager.all_results['total_cases'] > 0 else 0
    print(f"通过率: {pass_rate:.2f}%")
    print(f"总耗时: {test_manager.all_results['total_duration']:.2f}秒")
    
    # 计算转换时间统计
    conversion_times = []
    for test_case in test_manager.all_results["test_cases"]:
        if test_case.get("转换耗时"):
            conversion_times.append(test_case["转换耗时"])
    
    if conversion_times:
        avg_time = sum(conversion_times) / len(conversion_times)
        min_time = min(conversion_times)
        max_time = max(conversion_times)
        print(f"\n⏱️ 转换时间统计:")
        print(f"  平均转换时间: {avg_time:.2f}秒")
        print(f"  最短转换时间: {min_time:.2f}秒")
        print(f"  最长转换时间: {max_time:.2f}秒")

if __name__ == "__main__":
    main()