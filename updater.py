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

# ===================== 修复路径问题：区分临时资源和持久化目录 =====================
def get_exe_dir():
    """获取当前exe/脚本所在的永久目录（关键修复）"""
    if hasattr(sys, '_MEIPASS'):
        # 打包后：获取exe所在目录（而非临时目录）
        exe_path = os.path.dirname(sys.executable)
        return os.path.abspath(exe_path)
    else:
        # 开发模式：获取脚本所在目录
        return os.path.abspath(".")

def resource_path(relative_path):
    """仅用于读取打包进exe的资源文件（如图标等），不用于持久化文件"""
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ===================== 核心工具函数 =====================
def compare_versions(v1: str, v2: str) -> int:
    """
    比较两个版本号（x.y.z 格式）
    返回值：
    - 1: v1 > v2
    - 0: v1 == v2
    - -1: v1 < v2
    """
    def normalize_version(version: str) -> list[int]:
        parts = version.strip().split('.')
        return [int(part) if part.isdigit() else 0 for part in parts]
    
    v1_parts = normalize_version(v1)
    v2_parts = normalize_version(v2)
    
    # 补全版本号位数（如 1.2 和 1.2.0 视为相等）
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
    """计算文件哈希值，验证完整性"""
    hash_obj = hashlib.new(hash_algorithm)
    with open(file_path, 'rb') as f:
        while chunk := f.read(4096):
            hash_obj.update(chunk)
    return hash_obj.hexdigest()

# ===================== GUI主程序类 =====================
class AppUpdater(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("程序自动更新器")
        self.geometry("600x400")
        self.minsize(550, 380)
        
        # 获取exe/脚本所在的永久目录（核心修复）
        self.base_dir = get_exe_dir()
        
        # ===================== 配置项（已修改主程序路径）=====================
        self.config = {
            # 版本文件、解压目录、临时压缩包都放在exe所在目录
            "local_version_path": os.path.join(self.base_dir, "local_version.json"),
            "remote_version_url": "http://49.51.50.251/jiaoyu/version.json",
            "remote_archive_url": "http://49.51.50.251/jiaoyu/jiaoyu_win.zip",
            "archive_save_path": os.path.join(self.base_dir, "temp_update.zip"),
            "extract_dir": os.path.join(self.base_dir, "installed_program"),  # 解压到exe同目录
            # 核心修改：主程序路径改为 installed_program/jiaoyu_win/LuckyAi.exe
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
        
        # 构建UI
        self._build_ui()
        # 加载本地版本号
        self._load_local_version()
        
        # 启动时自动执行检查更新流程
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

    def _auto_update_flow(self):
        """自动更新流程：检查更新 -> 有更新则更新 -> 更新完成则运行程序"""
        # 先执行检查更新
        self._check_update_auto()

    def _check_update_auto(self):
        """自动检查更新（带后续流程）"""
        self.check_btn.config(state=tk.DISABLED)
        self._update_status("正在检查更新...")
        
        try:
            response = requests.get(self.config["remote_version_url"], timeout=10)
            response.raise_for_status()
            remote_data = response.json()
            self.remote_version = remote_data.get("version", "0.0.0")
            self.remote_version_label.config(text=self.remote_version)
            
            compare_result = compare_versions(self.remote_version, self.local_version)
            
            if compare_result > 0:
                self._update_status(f"发现新版本: {self.remote_version} (当前: {self.local_version})")
                self.update_btn.config(state=tk.NORMAL)
                # 有更新则自动执行更新
                self._update_thread_auto()
            elif compare_result == 0:
                self._update_status("当前已是最新版本")
                self.update_btn.config(state=tk.DISABLED)
                self.run_btn.config(state=tk.NORMAL)
                # 无更新则直接运行程序
                self._run_program()
            else:
                self._update_status(f"本地版本较新: {self.local_version} (远程: {self.remote_version})")
                self.update_btn.config(state=tk.DISABLED)
                self.run_btn.config(state=tk.NORMAL)
                # 本地版本更新则直接运行程序
                self._run_program()
                
        except Exception as e:
            self._update_status(f"检查更新失败: {str(e)}")
            messagebox.showerror("错误", f"检查更新失败:\n{str(e)}")
        finally:
            self.check_btn.config(state=tk.NORMAL)

    def _update_thread_auto(self):
        """自动更新线程（更新完成后自动运行程序）"""
        threading.Thread(target=self._perform_update_auto, daemon=True).start()

    def _perform_update_auto(self):
        """执行自动更新流程（更新完成后自动运行）"""
        self.update_btn.config(state=tk.DISABLED)
        self._update_status("开始下载更新包...")
        
        # 1. 下载压缩包
        if not self._download_archive():
            self._update_status("下载失败")
            self.check_btn.config(state=tk.NORMAL)
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
                messagebox.showerror("错误", "文件损坏，更新失败")
                self.check_btn.config(state=tk.NORMAL)
                return
        
        # 3. 解压文件
        self._update_status("开始解压文件...")
        if not self._extract_archive():
            self._update_status("解压失败")
            self.check_btn.config(state=tk.NORMAL)
            return
        
        # 4. 更新本地版本号
        self._save_local_version(self.remote_version)
        self._update_status("更新完成！")
        self.run_btn.config(state=tk.NORMAL)
        self.check_btn.config(state=tk.NORMAL)
        
        # 清理临时压缩包
        if os.path.exists(self.config["archive_save_path"]):
            os.remove(self.config["archive_save_path"])
        
        # 自动运行程序
        self._run_program()

    def _load_local_version(self):
        """加载本地版本号"""
        try:
            if os.path.exists(self.config["local_version_path"]):
                with open(self.config["local_version_path"], "r") as f:
                    data = json.load(f)
                    self.local_version = data.get("version", "0.0.0")
                self.local_version_label.config(text=self.local_version)
        except Exception as e:
            self._update_status(f"加载本地版本失败: {str(e)}")
            self.local_version_label.config(text="0.0.0")

    def _save_local_version(self, version: str):
        """保存本地版本号"""
        try:
            with open(self.config["local_version_path"], "w") as f:
                json.dump({"version": version}, f, indent=2)
            self.local_version = version
            self.local_version_label.config(text=version)
        except Exception as e:
            self._update_status(f"保存版本号失败: {str(e)}")

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
        """线程：手动检查更新（保留原有功能）"""
        threading.Thread(target=self._check_update, daemon=True).start()

    def _check_update(self):
        """手动检查远程版本并对比（保留原有功能）"""
        self.check_btn.config(state=tk.DISABLED)
        self._update_status("正在检查更新...")
        
        try:
            response = requests.get(self.config["remote_version_url"], timeout=10)
            response.raise_for_status()
            remote_data = response.json()
            self.remote_version = remote_data.get("version", "0.0.0")
            self.remote_version_label.config(text=self.remote_version)
            
            compare_result = compare_versions(self.remote_version, self.local_version)
            
            if compare_result > 0:
                self._update_status(f"发现新版本: {self.remote_version} (当前: {self.local_version})")
                self.update_btn.config(state=tk.NORMAL)
                self.run_btn.config(state=tk.DISABLED)
            elif compare_result == 0:
                self._update_status("当前已是最新版本")
                self.update_btn.config(state=tk.DISABLED)
                self.run_btn.config(state=tk.NORMAL)
            else:
                self._update_status(f"本地版本较新: {self.local_version} (远程: {self.remote_version})")
                self.update_btn.config(state=tk.DISABLED)
                self.run_btn.config(state=tk.NORMAL)
                
        except Exception as e:
            self._update_status(f"检查更新失败: {str(e)}")
            messagebox.showerror("错误", f"检查更新失败:\n{str(e)}")
        finally:
            self.check_btn.config(state=tk.NORMAL)

    def _update_thread(self):
        """线程：手动执行更新（保留原有功能）"""
        threading.Thread(target=self._perform_update, daemon=True).start()

    def _perform_update(self):
        """手动执行更新流程（保留原有功能）"""
        self.update_btn.config(state=tk.DISABLED)
        self._update_status("开始下载更新包...")
        
        # 1. 下载压缩包
        if not self._download_archive():
            self._update_status("下载失败")
            self.check_btn.config(state=tk.NORMAL)
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
                messagebox.showerror("错误", "文件损坏，更新失败")
                self.check_btn.config(state=tk.NORMAL)
                return
        
        # 3. 解压文件
        self._update_status("开始解压文件...")
        if not self._extract_archive():
            self._update_status("解压失败")
            self.check_btn.config(state=tk.NORMAL)
            return
        
        # 4. 更新本地版本号
        self._save_local_version(self.remote_version)
        self._update_status("更新完成！")
        self.run_btn.config(state=tk.NORMAL)
        self.check_btn.config(state=tk.NORMAL)
        
        # 清理临时压缩包
        if os.path.exists(self.config["archive_save_path"]):
            os.remove(self.config["archive_save_path"])

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
            return True
            
        except Exception as e:
            self._update_status(f"下载错误: {str(e)}")
            if os.path.exists(self.config["archive_save_path"]):
                os.remove(self.config["archive_save_path"])
            messagebox.showerror("错误", f"下载失败:\n{str(e)}")
            return False

    def _extract_archive(self) -> bool:
        """解压压缩包"""
        try:
            # 清空原有解压目录
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
            return True
            
        except Exception as e:
            self._update_status(f"解压错误: {str(e)}")
            messagebox.showerror("错误", f"解压失败:\n{str(e)}")
            return False

    def _run_program(self):
        """运行主程序"""
        import subprocess
        
        try:
            program_path = self.config["main_program_path"]
            if not os.path.exists(program_path):
                raise FileNotFoundError(f"程序文件不存在: {program_path}")
            
            self._update_status("正在启动程序...")
            # 切换到程序所在目录运行，避免路径问题
            subprocess.Popen([program_path], cwd=os.path.dirname(program_path))
            self._update_status("程序已启动")
            
        except Exception as e:
            self._update_status(f"启动失败: {str(e)}")
            messagebox.showerror("错误", f"启动程序失败:\n{str(e)}")

# ===================== 程序入口 =====================
def main():
    if sys.version_info < (3, 6):
        print("错误：需要Python 3.6或更高版本")
        sys.exit(1)
    
    try:
        import requests
    except ImportError:
        print("正在安装依赖库 requests...")
        os.system(f"{sys.executable} -m pip install requests")
        import requests
    
    app = AppUpdater()
    app.mainloop()

if __name__ == "__main__":
    main()