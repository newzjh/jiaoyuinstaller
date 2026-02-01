import os
import sys
import shutil
import json
import threading
import requests
import zipfile
import tarfile
from pathlib import Path
import hashlib
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Tuple
import logging
import time

# Windows平台特定导入
try:
    import win32com.client
    import winreg
    import pythoncom
    WINDOWS = True
except ImportError:
    WINDOWS = False

# ===================== 基础配置与工具函数 =====================
def get_exe_dir():
    """获取当前exe/脚本所在的永久目录"""
    if hasattr(sys, '_MEIPASS'):
        exe_path = os.path.dirname(sys.executable)
        return os.path.abspath(exe_path)
    else:
        return os.path.abspath(".")

# 初始化日志
BASE_DIR = get_exe_dir()
logging.basicConfig(
    filename=os.path.join(BASE_DIR, "update_log.log"),
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def resource_path(relative_path):
    """仅用于读取打包进exe的资源文件"""
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def compare_versions(v1: str, v2: str) -> int:
    """比较版本号：1(v1>v2) / 0(相等) / -1(v1<v2)"""
    def normalize_version(version: str) -> list[int]:
        parts = version.strip().split('.')
        return [int(part) if part.isdigit() else 0 for part in parts]
    
    v1_parts = normalize_version(v1)
    v2_parts = normalize_version(v2)
    max_len = max(len(v1_parts), len(v2_parts))
    v1_parts += [0] * (max_len - len(v1_parts))
    v2_parts += [0] * (max_len - len(v2_parts))
    
    for p1, p2 in zip(v1_parts, v2_parts):
        if p1 > p2:
            return 1
        elif p1 < p2:
            return -1
    return 0

def calculate_file_hash(file_path: str, hash_algorithm: str = 'md5') -> str:
    """计算文件哈希值"""
    if not os.path.exists(file_path):
        return ""
    hash_obj = hashlib.new(hash_algorithm)
    with open(file_path, 'rb') as f:
        while chunk := f.read(4096):
            hash_obj.update(chunk)
    return hash_obj.hexdigest()

def retry(max_retries=3, delay=1):
    """重试装饰器"""
    def decorator(func):
        def wrapper(*args, **kwargs):
            for i in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    logging.warning(f"第{i+1}次执行失败: {str(e)}，重试中...")
                    if i == max_retries - 1:
                        raise
                    time.sleep(delay)
            return None
        return wrapper
    return decorator

def create_shortcut(target_path, shortcut_path, description="", icon_path=None):
    """创建Windows快捷方式
    
    Args:
        target_path: 目标文件路径
        shortcut_path: 快捷方式保存路径
        description: 快捷方式描述
        icon_path: 图标路径
    """
    if not WINDOWS:
        return False
    
    try:
        # 初始化COM库
        pythoncom.CoInitialize()
        
        # 确保目标目录存在
        shortcut_dir = os.path.dirname(shortcut_path)
        if shortcut_dir and not os.path.exists(shortcut_dir):
            try:
                os.makedirs(shortcut_dir, exist_ok=True)
                logging.info(f"创建目录成功: {shortcut_dir}")
            except Exception as dir_error:
                logging.warning(f"创建目录失败: {str(dir_error)}，将尝试直接创建快捷方式")
        
        shell = win32com.client.Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = target_path
        shortcut.WorkingDirectory = os.path.dirname(target_path)
        if description:
            shortcut.Description = description
        if icon_path:
            shortcut.IconLocation = icon_path
        shortcut.Save()
        logging.info(f"快捷方式创建成功: {shortcut_path}")
        return True
    except PermissionError:
        logging.warning(f"创建快捷方式无权限: {shortcut_path}，可能需要管理员权限或被安全软件阻止")
        return False
    except Exception as e:
        logging.error(f"创建快捷方式失败: {str(e)}")
        return False
    finally:
        # 释放COM库
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def get_desktop_path():
    """获取桌面路径"""
    if not WINDOWS:
        return ""
    
    try:
        # 初始化COM库
        pythoncom.CoInitialize()
        
        shell = win32com.client.Dispatch('WScript.Shell')
        result = shell.SpecialFolders('Desktop')
        return result
    except Exception as e:
        logging.error(f"获取桌面路径失败: {str(e)}")
        return os.path.join(os.path.expanduser('~'), 'Desktop')
    finally:
        # 释放COM库
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def get_start_menu_path():
    """获取开始菜单路径"""
    if not WINDOWS:
        return ""
    
    try:
        # 获取当前用户的开始菜单程序目录
        start_menu = os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', 'Microsoft', 'Windows', 'Start Menu', 'Programs')
        return start_menu
    except Exception as e:
        logging.error(f"获取开始菜单路径失败: {str(e)}")
        return ""

# ===================== GUI主程序类 =====================
class AppUpdater(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("LuckyAI程序自动更新器")
        self.geometry("600x400")
        self.minsize(550, 380)
        
        # 核心状态控制
        self.is_auto_running = False  # 标记自动流程是否在运行
        self.base_dir = BASE_DIR
        
        # 配置项
        self.config = {
            "local_version_path": os.path.join(self.base_dir, "local_version.json"),
            "remote_version_url": "http://49.51.50.251/jiaoyu/version.json",
            "remote_archive_url": "http://49.51.50.251/jiaoyu/jiaoyu_win.zip",
            "archive_save_path": os.path.join(self.base_dir, "temp_update.zip"),
            "extract_dir": os.path.join(self.base_dir, "installed_program"),
            "main_program_path": os.path.join(self.base_dir, "installed_program/jiaoyu_win/LuckyAi.exe"),
            "expected_hash": None,
            "hash_algorithm": "md5"
        }
        
        # 初始化变量
        self.local_version = "0.0.0"
        self.remote_version = "0.0.0"
        self.download_progress = tk.DoubleVar()
        self.extract_progress = tk.DoubleVar()
        self.status_text = tk.StringVar(value="就绪")
        self.main_program_missing = False
        
        # 构建UI
        self._build_ui()
        # 初始化流程
        self._load_local_version()
        self._check_main_program_exists()
        
        # 启动自动流程
        self.after(100, self._auto_update_flow)

    def _build_ui(self):
        """构建GUI界面"""
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        padding = {"padx": 10, "pady": 8}
        
        # 版本信息区域
        version_frame = ttk.LabelFrame(self, text="版本信息", padding=10)
        version_frame.grid(row=0, column=0, sticky="nsew", **padding)
        version_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(version_frame, text="本地版本:", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=3)
        self.local_version_label = ttk.Label(version_frame, text=self.local_version, font=("Arial", 10))
        self.local_version_label.grid(row=0, column=1, sticky=tk.W, pady=3)
        
        ttk.Label(version_frame, text="远程版本:", font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, pady=3)
        self.remote_version_label = ttk.Label(version_frame, text=self.remote_version, font=("Arial", 10))
        self.remote_version_label.grid(row=1, column=1, sticky=tk.W, pady=3)
        
        # 下载进度区域
        download_frame = ttk.LabelFrame(self, text="下载进度", padding=10)
        download_frame.grid(row=1, column=0, sticky="nsew", **padding)
        download_frame.grid_columnconfigure(0, weight=1)
        
        self.download_bar = ttk.Progressbar(
            download_frame, variable=self.download_progress, maximum=100
        )
        self.download_bar.grid(row=0, column=0, sticky="ew", pady=5)
        self.download_label = ttk.Label(download_frame, text="0%")
        self.download_label.grid(row=1, column=0, pady=2)
        
        # 解压进度区域
        extract_frame = ttk.LabelFrame(self, text="解压进度", padding=10)
        extract_frame.grid(row=2, column=0, sticky="nsew", **padding)
        extract_frame.grid_columnconfigure(0, weight=1)
        
        self.extract_bar = ttk.Progressbar(
            extract_frame, variable=self.extract_progress, maximum=100
        )
        self.extract_bar.grid(row=0, column=0, sticky="ew", pady=5)
        self.extract_label = ttk.Label(extract_frame, text="0%")
        self.extract_label.grid(row=1, column=0, pady=2)
        
        # 状态和按钮区域
        bottom_frame = ttk.Frame(self)
        bottom_frame.grid(row=3, column=0, sticky="nsew", **padding)
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_columnconfigure(1, weight=0)
        
        status_label = ttk.Label(bottom_frame, textvariable=self.status_text, font=("Arial", 9))
        status_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        button_frame = ttk.Frame(bottom_frame)
        button_frame.grid(row=0, column=1, sticky=tk.E, padx=5, pady=5)
        
        self.check_btn = ttk.Button(
            button_frame, text="检查更新", command=self._check_update_thread, width=10
        )
        self.check_btn.grid(row=0, column=0, padx=3, pady=2)
        
        self.update_btn = ttk.Button(
            button_frame, text="立即更新", command=self._update_thread, state=tk.DISABLED, width=10
        )
        self.update_btn.grid(row=0, column=1, padx=3, pady=2)
        
        self.run_btn = ttk.Button(
            button_frame, text="运行程序", command=self._run_program, state=tk.DISABLED, width=10
        )
        self.run_btn.grid(row=0, column=2, padx=3, pady=2)

    def _disable_all_buttons(self):
        """禁用所有操作按钮"""
        self.check_btn.config(state=tk.DISABLED)
        self.update_btn.config(state=tk.DISABLED)
        self.run_btn.config(state=tk.DISABLED)
        self.update_idletasks()  # 强制刷新UI

    def _enable_buttons_normal(self):
        """恢复按钮正常状态"""
        self.check_btn.config(state=tk.NORMAL)
        # 根据状态决定更新/运行按钮
        if compare_versions(self.remote_version, self.local_version) > 0 or self.main_program_missing:
            self.update_btn.config(state=tk.NORMAL)
            self.run_btn.config(state=tk.DISABLED)
        else:
            self.update_btn.config(state=tk.DISABLED)
            self.run_btn.config(state=tk.NORMAL)
        self.update_idletasks()

    def _check_main_program_exists(self):
        """检查主程序是否存在"""
        self.main_program_missing = not os.path.exists(self.config["main_program_path"])
        if self.main_program_missing:
            self._update_status(f"主程序文件缺失: {self.config['main_program_path']}")
            logging.warning(f"主程序文件缺失: {self.config['main_program_path']}")
        else:
            self._update_status("主程序文件存在")
            logging.info("主程序文件存在")

    def _auto_update_flow(self):
        """自动更新流程（核心调整：仅获取远程版本号，不提前同步）"""
        self.is_auto_running = True
        self._disable_all_buttons()  # 自动流程开始就禁用所有按钮
        
        try:
            # 第一步：仅获取远程版本号并刷新UI（不同步本地版本号）
            self._update_status("正在获取远程版本信息...")
            self._fetch_remote_version()
            
            # 第二步：判断是否需要修复/更新
            if self.main_program_missing:
                self._update_status("主程序缺失，开始修复下载...")
                self._update_thread_auto(fix_mode=True)
            else:
                self._check_update_auto()
        except Exception as e:
            logging.error(f"自动流程初始化失败: {str(e)}")
            self._update_status(f"初始化失败: {str(e)}")
            self._enable_buttons_normal()
            self.is_auto_running = False

    @retry(max_retries=3, delay=1)
    def _fetch_remote_version(self):
        """单独获取远程版本号（仅刷新UI，不修改本地版本）"""
        try:
            response = requests.get(self.config["remote_version_url"], timeout=10)
            response.raise_for_status()
            remote_data = response.json()
            self.remote_version = remote_data.get("version", "0.0.0")
            
            # 仅刷新远程版本号UI，不修改本地版本
            self.remote_version_label.config(text=self.remote_version)
            self.update_idletasks()
            logging.info(f"远程版本号获取成功：{self.remote_version}，本地版本号仍为：{self.local_version}")
        except Exception as e:
            logging.error(f"获取远程版本号失败: {str(e)}")
            raise

    def _check_update_auto(self):
        """自动检查版本更新"""
        self._update_status("正在检查更新...")
        
        try:
            compare_result = compare_versions(self.remote_version, self.local_version)
            logging.info(f"版本对比：本地{self.local_version}，远程{self.remote_version}，结果{compare_result}")
            
            if compare_result > 0:
                self._update_status(f"发现新版本: {self.remote_version} (当前: {self.local_version})")
                self._update_thread_auto()
            elif compare_result == 0:
                self._update_status("当前已是最新版本")
                self._enable_buttons_normal()
                self._run_program()
            else:
                self._update_status(f"本地版本较新: {self.local_version} (远程: {self.remote_version})")
                self._enable_buttons_normal()
                self._run_program()
                
        except Exception as e:
            error_msg = f"检查更新失败: {str(e)}"
            self._update_status(error_msg)
            logging.error(error_msg)
            messagebox.showerror("错误", f"检查更新失败:\n{str(e)}")
            self._enable_buttons_normal()
            self.is_auto_running = False

    def _update_thread_auto(self, fix_mode=False):
        """自动更新线程"""
        # 子线程中执行下载解压，避免阻塞UI
        def worker():
            try:
                self._perform_update_auto(fix_mode)
            finally:
                # 恢复按钮状态+标记流程结束
                self.is_auto_running = False
                self._enable_buttons_normal()
        
        threading.Thread(target=worker, daemon=True).start()

    def _perform_update_auto(self, fix_mode=False):
        """执行自动更新/修复流程（核心调整：仅在解压完成后同步版本号）"""
        status_text = "开始修复下载主程序..." if fix_mode else "开始下载更新包..."
        self._update_status(status_text)
        
        # 1. 下载压缩包
        if not self._download_archive():
            self._update_status("下载失败")
            return
        
        # 2. 验证文件完整性
        if self.config["expected_hash"]:
            self._update_status("验证文件完整性...")
            file_hash = calculate_file_hash(
                self.config["archive_save_path"],
                self.config["hash_algorithm"]
            )
            if file_hash != self.config["expected_hash"]:
                self._update_status("文件哈希值不匹配")
                os.remove(self.config["archive_save_path"])
                messagebox.showerror("错误", "文件损坏，更新/修复失败")
                logging.error(f"哈希校验失败：预期{self.config['expected_hash']}，实际{file_hash}")
                return
        
        # 3. 解压文件
        self._update_status("开始解压文件...")
        if not self._extract_archive():
            self._update_status("解压失败")
            return
        
        # 4. 核心调整：仅在下载解压完成后，同步本地版本号（关键时机）
        self._update_status(f"同步本地版本号为 {self.remote_version}...")
        self._save_local_version(self.remote_version)  # 写入版本文件
        self.local_version = self.remote_version  # 更新内存中的本地版本号
        self.local_version_label.config(text=self.local_version)  # 刷新本地版本UI
        self.update_idletasks()
        
        # 5. 清理临时文件
        if os.path.exists(self.config["archive_save_path"]):
            os.remove(self.config["archive_save_path"])
        
        # 6. 完成提示并运行程序
        finish_text = "主程序修复完成！" if fix_mode else "更新完成！"
        self._update_status(finish_text)
        logging.info(f"{finish_text} 本地版本号已同步为：{self.local_version}（仅在解压完成后同步）")
        
        # 重新检查主程序
        self._check_main_program_exists()
        # 恢复按钮+运行程序
        self._enable_buttons_normal()
        self._run_program()

    def _load_local_version(self):
        """加载本地版本号（仅读取，不提前同步）"""
        try:
            if os.path.exists(self.config["local_version_path"]):
                with open(self.config["local_version_path"], "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.local_version = data.get("version", "0.0.0")
                self.local_version_label.config(text=self.local_version)
                logging.info(f"加载本地版本号：{self.local_version}")
            else:
                self._update_status("本地版本文件缺失，初始化版本信息为0.0.0...")
                self.local_version = "0.0.0"
                self._save_local_version(self.local_version)  # 仅初始化，不同步远程
                self.local_version_label.config(text=self.local_version)
                logging.info("本地版本文件缺失，已初始化为0.0.0（未同步远程）")
        except Exception as e:
            error_msg = f"加载本地版本失败: {str(e)}"
            self._update_status(error_msg)
            logging.error(error_msg)
            # 加载失败时仅初始化，不同步远程
            self.local_version = "0.0.0"
            self._save_local_version(self.local_version)
            self.local_version_label.config(text="0.0.0")

    def _save_local_version(self, version: str):
        """保存本地版本号（增强可靠性，仅在解压完成后调用）"""
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(self.config["local_version_path"]), exist_ok=True)
            # 写入版本文件（指定编码，避免乱码）
            with open(self.config["local_version_path"], "w", encoding="utf-8") as f:
                json.dump({"version": version}, f, indent=2, ensure_ascii=False)
            logging.info(f"本地版本号已写入文件：{version}，路径：{self.config['local_version_path']}")
        except Exception as e:
            error_msg = f"保存版本号失败: {str(e)}"
            self._update_status(error_msg)
            logging.error(error_msg)
            # 保存失败时弹窗提示
            messagebox.warning("警告", f"版本号保存失败:\n{str(e)}\n可能导致后续更新异常！")

    def _update_status(self, text: str):
        """更新状态文本"""
        self.status_text.set(text)
        self.update_idletasks()

    def _update_download_progress(self, progress: float):
        """更新下载进度条"""
        self.download_progress.set(progress)
        self.download_label.config(text=f"{progress:.1f}%")
        self.update_idletasks()

    def _update_extract_progress(self, progress: float):
        """更新解压进度条"""
        self.extract_progress.set(progress)
        self.extract_label.config(text=f"{progress:.1f}%")
        self.update_idletasks()

    def _check_update_thread(self):
        """手动检查更新线程"""
        if self.is_auto_running:
            messagebox.showinfo("提示", "自动更新流程正在运行，请稍后再试！")
            return
        
        threading.Thread(target=self._check_update, daemon=True).start()

    def _check_update(self):
        """手动检查远程版本（仅获取，不提前同步）"""
        self._disable_all_buttons()
        self._update_status("正在检查更新...")
        
        try:
            # 先获取最新版本号（仅刷新UI，不修改本地版本）
            self._fetch_remote_version()
            
            compare_result = compare_versions(self.remote_version, self.local_version)
            
            if compare_result > 0:
                self._update_status(f"发现新版本: {self.remote_version} (当前: {self.local_version})")
                self.update_btn.config(state=tk.NORMAL)
                self.run_btn.config(state=tk.DISABLED)
            elif compare_result == 0:
                # 检查主程序是否缺失
                self._check_main_program_exists()
                if self.main_program_missing:
                    self._update_status("版本最新但主程序缺失，请点击【立即更新】修复")
                    self.update_btn.config(state=tk.NORMAL)
                    self.run_btn.config(state=tk.DISABLED)
                else:
                    self._update_status("当前已是最新版本，主程序正常")
                    self.update_btn.config(state=tk.DISABLED)
                    self.run_btn.config(state=tk.NORMAL)
            else:
                self._update_status(f"本地版本较新: {self.local_version} (远程: {self.remote_version})")
                self.update_btn.config(state=tk.DISABLED)
                self.run_btn.config(state=tk.NORMAL)
                
        except Exception as e:
            error_msg = f"检查更新失败: {str(e)}"
            self._update_status(error_msg)
            logging.error(error_msg)
            messagebox.showerror("错误", f"检查更新失败:\n{str(e)}")
        finally:
            self._enable_buttons_normal()

    def _update_thread(self):
        """手动更新线程"""
        if self.is_auto_running:
            messagebox.showinfo("提示", "自动更新流程正在运行，请稍后再试！")
            return
        
        fix_mode = self.main_program_missing and (compare_versions(self.remote_version, self.local_version) == 0)
        threading.Thread(target=self._perform_update_manual, args=(fix_mode,), daemon=True).start()

    def _perform_update_manual(self, fix_mode=False):
        """手动执行更新/修复流程（同步逻辑与自动流程一致）"""
        self._disable_all_buttons()
        status_text = "开始修复下载主程序..." if fix_mode else "开始下载更新包..."
        self._update_status(status_text)
        
        try:
            # 1. 下载压缩包
            if not self._download_archive():
                self._update_status("下载失败")
                return
            
            # 2. 验证文件完整性
            if self.config["expected_hash"]:
                self._update_status("验证文件完整性...")
                file_hash = calculate_file_hash(
                    self.config["archive_save_path"],
                    self.config["hash_algorithm"]
                )
                if file_hash != self.config["expected_hash"]:
                    self._update_status("文件哈希值不匹配")
                    os.remove(self.config["archive_save_path"])
                    messagebox.showerror("错误", "文件损坏，更新/修复失败")
                    logging.error(f"哈希校验失败：预期{self.config['expected_hash']}，实际{file_hash}")
                    return
            
            # 3. 解压文件
            self._update_status("开始解压文件...")
            if not self._extract_archive():
                self._update_status("解压失败")
                return
            
            # 4. 仅在解压完成后同步版本号
            self._update_status(f"同步本地版本号为 {self.remote_version}...")
            self._save_local_version(self.remote_version)
            self.local_version = self.remote_version
            self.local_version_label.config(text=self.local_version)
            self.update_idletasks()
            
            # 5. 清理临时文件
            if os.path.exists(self.config["archive_save_path"]):
                os.remove(self.config["archive_save_path"])
            
            # 6. 完成提示
            finish_text = "主程序修复完成！" if fix_mode else "更新完成！"
            self._update_status(finish_text)
            logging.info(f"{finish_text} 本地版本号已同步为：{self.local_version}（手动操作，解压完成后同步）")
            
            # 重新检查主程序
            self._check_main_program_exists()
        finally:
            self._enable_buttons_normal()

    def _download_archive(self) -> bool:
        """下载压缩包（带进度）"""
        try:
            response = requests.get(
                self.config["remote_archive_url"],
                stream=True,
                timeout=30
            )
            response.raise_for_status()
            
            total_size = int(response.headers.get("Content-Length", 0))
            downloaded_size = 0
            
            with open(self.config["archive_save_path"], "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded_size += len(chunk)
                        if total_size > 0:
                            progress = (downloaded_size / total_size) * 100
                            self._update_download_progress(progress)
            
            self._update_download_progress(100.0)
            self._update_status("下载完成")
            logging.info(f"更新包下载完成，路径：{self.config['archive_save_path']}，大小：{total_size/1024/1024:.2f}MB")
            return True
            
        except Exception as e:
            error_msg = f"下载错误: {str(e)}"
            self._update_status(error_msg)
            logging.error(error_msg)
            if os.path.exists(self.config["archive_save_path"]):
                os.remove(self.config["archive_save_path"])
            messagebox.showerror("错误", f"下载失败:\n{str(e)}")
            return False

    def _extract_archive(self) -> bool:
        """解压压缩包"""
        try:
            if os.path.exists(self.config["extract_dir"]):
                shutil.rmtree(self.config["extract_dir"])
            os.makedirs(self.config["extract_dir"], exist_ok=True)
            
            archive_path = self.config["archive_save_path"]
            if archive_path.endswith('.zip'):
                with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                    total_files = len(zip_ref.infolist())
                    extracted_files = 0
                    
                    for file in zip_ref.infolist():
                        zip_ref.extract(file, self.config["extract_dir"])
                        extracted_files += 1
                        progress = (extracted_files / total_files) * 100
                        self._update_extract_progress(progress)
            else:
                raise ValueError(f"不支持的格式: {archive_path}")
            
            self._update_extract_progress(100.0)
            self._update_status("解压完成")
            logging.info(f"更新包解压完成，路径：{self.config['extract_dir']}")
            return True
            
        except Exception as e:
            error_msg = f"解压错误: {str(e)}"
            self._update_status(error_msg)
            logging.error(error_msg)
            messagebox.showerror("错误", f"解压失败:\n{str(e)}")
            return False

    def _run_program(self):
        """运行主程序"""
        import subprocess
        
        try:
            program_path = self.config["main_program_path"]
            if not os.path.exists(program_path):
                raise FileNotFoundError(f"程序文件不存在: {program_path}")
            
            # 在Windows平台下创建快捷方式
            if WINDOWS:
                self._create_shortcuts(program_path)
            
            subprocess.Popen([program_path], cwd=os.path.dirname(program_path))
            self._update_status("主程序已启动")
            logging.info(f"主程序启动成功，路径：{program_path}")
            self.quit()
        except Exception as e:
            error_msg = f"启动程序失败: {str(e)}"
            self._update_status(error_msg)
            logging.error(error_msg)
            messagebox.showerror("错误", f"启动程序失败:\n{str(e)}")
    
    def _create_shortcuts(self, program_path):
        """创建桌面快捷方式和开始菜单快捷方式"""
        if not WINDOWS:
            return
        
        try:
            # 获取快捷方式名称
            shortcut_name = "LuckyAI.lnk"
            description = "LuckyAI登录器"
            
            # 获取图标路径（使用程序自身作为图标）
            icon_path = program_path
            
            # 创建桌面快捷方式
            desktop_path = get_desktop_path()
            if desktop_path:
                desktop_shortcut = os.path.join(desktop_path, shortcut_name)
                if not os.path.exists(desktop_shortcut):
                    success = create_shortcut(program_path, desktop_shortcut, description, icon_path)
                    if not success:
                        logging.warning("桌面快捷方式创建失败，但不影响程序运行")
                else:
                    logging.info("桌面快捷方式已存在，跳过创建")
            
            # 创建开始菜单快捷方式
            start_menu_path = get_start_menu_path()
            if start_menu_path:
                start_menu_shortcut = os.path.join(start_menu_path, shortcut_name)
                if not os.path.exists(start_menu_shortcut):
                    success = create_shortcut(program_path, start_menu_shortcut, description, icon_path)
                    if not success:
                        logging.warning("开始菜单快捷方式创建失败，但不影响程序运行")
                else:
                    logging.info("开始菜单快捷方式已存在，跳过创建")
            
            logging.info("快捷方式创建流程完成")
        except Exception as e:
            logging.error(f"创建快捷方式过程中出错: {str(e)}")
            # 即使出错也不影响主程序运行
            pass

# ===================== 程序入口 =====================
if __name__ == "__main__":
    if sys.version_info < (3, 6):
        messagebox.showerror("错误", "需要Python 3.6或更高版本")
        sys.exit(1)
    
    try:
        import requests
    except ImportError:
        messagebox.showinfo("提示", "正在安装依赖库 requests...")
        os.system(f"{sys.executable} -m pip install requests")
        import requests
    
    # 检查并安装pywin32库（Windows平台需要）
    if WINDOWS:
        try:
            import win32com.client
            import winreg
        except ImportError:
            messagebox.showinfo("提示", "正在安装依赖库 pywin32...")
            os.system(f"{sys.executable} -m pip install pywin32")
            try:
                import win32com.client
                import winreg
                WINDOWS = True
            except ImportError:
                WINDOWS = False
                logging.warning("安装pywin32失败，将无法创建快捷方式")
    
    app = AppUpdater()
    app.mainloop()