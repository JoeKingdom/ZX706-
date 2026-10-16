import os
import re
import sys
import json
import subprocess
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# 尝试导入第三方库（用于读取音频标签和生成Excel）
try:
    import mutagen
    HAS_MUTAGEN = True
except ImportError:
    HAS_MUTAGEN = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# 支持的音乐文件后缀（包含索尼 Walkman 常用的 Hi-Res 格式如 dsf/dff/flac/wav/m4a 等）
AUDIO_EXTENSIONS = {
    '.mp3', '.flac', '.wav', '.aac', '.m4a', '.dsf', '.dff',
    '.ape', '.ogg', '.wma', '.alac', '.aiff', '.aif', '.opus'
}


class ZX706MusicExporterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("索尼 ZX706 音乐库导出备份工具")
        self.root.geometry("760x580")
        self.root.minsize(700, 520)

        self.is_running = False
        self.last_output_dir = ""

        self._build_ui()
        self._check_dependencies()

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 顶部说明区域
        title_label = ttk.Label(
            main_frame,
            text="🎵 索尼 Walkman (ZX706) 歌曲清单备份工具",
            font=("Microsoft YaHei", 14, "bold")
        )
        title_label.pack(anchor=tk.W, pady=(0, 8))

        desc_label = ttk.Label(
            main_frame,
            text="支持格式：MP3 / FLAC / WAV / AAC / M4A / DSF / DFF / APE / OGG 等全格式音频\n"
                 "导出结果：自动在桌面（或指定位置）新建备份文件夹，并生成 Excel 表格 (.xlsx) 与纯文本清单 (.txt)",
            foreground="#555555"
        )
        desc_label.pack(anchor=tk.W, pady=(0, 12))

        # 2. 扫描源设置区域
        source_frame = ttk.LabelFrame(main_frame, text=" 第一步：选择音乐扫描方式 ", padding="10")
        source_frame.pack(fill=tk.X, pady=(0, 10))

        # 方式 A：标准文件夹路径
        folder_row = ttk.Frame(source_frame)
        folder_row.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(folder_row, text="文件夹路径：").pack(side=tk.LEFT)
        self.folder_var = tk.StringVar()
        self.folder_entry = ttk.Entry(folder_row, textvariable=self.folder_var)
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        ttk.Button(
            folder_row, text="📂 浏览选择文件夹", command=self.select_folder
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(
            folder_row, text="🚀 扫描该文件夹并导出", command=self.start_folder_scan
        ).pack(side=tk.LEFT, padx=2)

        # 分隔提示与 方式 B：USB MTP 直接扫描
        mtp_row = ttk.Frame(source_frame)
        mtp_row.pack(fill=tk.X, pady=(4, 0))

        ttk.Label(
            mtp_row,
            text="USB直连模式（无需盘符，直接读取『此电脑』中的 ZX706 便携设备）：",
            foreground="#0055aa"
        ).pack(side=tk.LEFT)

        ttk.Button(
            mtp_row,
            text="📱 直接扫描 USB 连接的 ZX706 设备",
            command=self.start_mtp_scan
        ).pack(side=tk.RIGHT, padx=2)

        # 3. 导出与命名解析设置区域
        setting_frame = ttk.LabelFrame(main_frame, text=" 第二步：解析与导出设置 ", padding="10")
        setting_frame.pack(fill=tk.X, pady=(0, 10))

        # 文件名解析规则（当音频无内嵌标签或使用MTP模式时生效）
        rule_row = ttk.Frame(setting_frame)
        rule_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(rule_row, text="当文件名含『-』且无内嵌标签时，原文件名格式为：").pack(side=tk.LEFT)

        self.filename_pattern_var = tk.StringVar(value="artist_first")
        ttk.Radiobutton(
            rule_row, text="歌手 - 歌名 (如：周杰伦 - 晴天.flac)",
            variable=self.filename_pattern_var, value="artist_first"
        ).pack(side=tk.LEFT, padx=8)
        ttk.Radiobutton(
            rule_row, text="歌名 - 歌手 (如：晴天 - 周杰伦.flac)",
            variable=self.filename_pattern_var, value="title_first"
        ).pack(side=tk.LEFT, padx=8)

        # 导出保存位置
        out_row = ttk.Frame(setting_frame)
        out_row.pack(fill=tk.X)
        ttk.Label(out_row, text="导出保存目录：").pack(side=tk.LEFT)

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        default_out = os.path.join(desktop_path, "ZX706_歌单Excel备份")
        self.output_dir_var = tk.StringVar(value=default_out)

        ttk.Entry(out_row, textvariable=self.output_dir_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=5
        )
        ttk.Button(
            out_row, text="更改保存位置", command=self.select_output_dir
        ).pack(side=tk.LEFT, padx=2)
        self.open_dir_btn = ttk.Button(
            out_row, text="📁 打开导出文件夹", command=self.open_output_dir
        )
        self.open_dir_btn.pack(side=tk.LEFT, padx=2)

        # 4. 进度条与日志区域
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(0, 5))

        self.status_var = tk.StringVar(value="准备就绪，请选择扫描方式。")
        ttk.Label(progress_frame, textvariable=self.status_var).pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(main_frame, mode="determinate")
        self.progress.pack(fill=tk.X, pady=(0, 8))

        self.log_area = scrolledtext.ScrolledText(
            main_frame, height=12, font=("Consolas", 9), state=tk.DISABLED
        )
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def _check_dependencies(self):
        missing = []
        if not HAS_OPENPYXL:
            missing.append("openpyxl")
        if not HAS_MUTAGEN:
            missing.append("mutagen")

        if missing:
            cmd = f"pip install {' '.join(missing)}"
            self.log(f"⚠️ 提示：检测到未安装扩展库 {', '.join(missing)}")
            self.log(f"👉 建议在命令行运行：{cmd}")
            self.log("   - 安装 openpyxl 后可导出标准 .xlsx Excel表格（未安装则自动降级导出可用Excel打开的 .csv 文件）")
            self.log("   - 安装 mutagen 后可读取音频文件内部歌手/歌名标签（未安装则直接通过文件名提取）\n")
        else:
            self.log("✅ 环境检查通过：已加载 mutagen（音频标签解析）与 openpyxl（Excel导出）。\n")

    def log(self, message):
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)

    def select_folder(self):
        folder = filedialog.askdirectory(title="选择音乐所在文件夹（如TF卡或本地目录）")
        if folder:
            self.folder_var.set(folder)

    def select_output_dir(self):
        folder = filedialog.askdirectory(title="选择 Excel 备份文件夹的存放位置")
        if folder:
            target = os.path.join(folder, "ZX706_歌单Excel备份")
            self.output_dir_var.set(target)

    def open_output_dir(self):
        out_dir = self.output_dir_var.get().strip()
        if not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(out_dir)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", out_dir])
        else:
            subprocess.Popen(["xdg-open", out_dir])

    @staticmethod
    def parse_filename(filename, pattern_mode="artist_first"):
        """从文件名中智能提取 歌名 和 歌手"""
        name_body = os.path.splitext(filename)[0].strip()
        # 去除常见的开头曲目编号，例如 "01. ", "01 - ", "[01] ", "01_"
        name_body = re.sub(r'^\s*[\[\(]?\d{1,3}[\]\)]?[\.\-_~\s]+\s*', '', name_body)

        if " - " in name_body:
            parts = name_body.split(" - ", 1)
        elif "-" in name_body:
            parts = name_body.split("-", 1)
        else:
            return name_body.strip(), "未知歌手"

        part1, part2 = parts[0].strip(), parts[1].strip()
        if pattern_mode == "artist_first":
            artist, title = part1, part2
        else:
            title, artist = part1, part2

        return (title or name_body), (artist or "未知歌手")

    def extract_song_metadata(self, filepath, filename, relative_folder):
        """优先读取音频内部标签，若缺失则回退到文件名解析"""
        ext = os.path.splitext(filename)[1].lower()
        title, artist, album = "", "", ""

        if HAS_MUTAGEN and filepath and os.path.exists(filepath):
            try:
                audio = mutagen.File(filepath, easy=True)
                if audio:
                    title = str(audio.get("title", [""])[0]).strip()
                    artist = str(audio.get("artist", [""])[0]).strip()
                    album = str(audio.get("album", [""])[0]).strip()
                else:
                    # 针对 DSF 等特殊格式尝试读取原始 ID3 标签
                    raw_audio = mutagen.File(filepath)
                    if raw_audio and raw_audio.tags:
                        tags = raw_audio.tags
                        if "TIT2" in tags:
                            title = str(tags["TIT2"]).strip()
                        if "TPE1" in tags:
                            artist = str(tags["TPE1"]).strip()
                        if "TALB" in tags:
                            album = str(tags["TALB"]).strip()
            except Exception:
                pass

        # 如果标签中没有歌名或歌手，则使用文件名补充
        fn_title, fn_artist = self.parse_filename(filename, self.filename_pattern_var.get())
        if not title:
            title = fn_title
        if not artist:
            artist = fn_artist

        combined_text = f"{title} - {artist}" if artist != "未知歌手" else f"{title} - 未知歌手"

        return {
            "combined": combined_text,
            "title": title,
            "artist": artist,
            "album": album,
            "ext": ext,
            "filename": filename,
            "folder": relative_folder
        }

    def start_folder_scan(self):
        if self.is_running:
            return
        folder = self.folder_var.get().strip()
        if not folder or not os.path.exists(folder):
            messagebox.showwarning("提示", "请先选择一个有效的音乐文件夹路径！\n\n如果是直接用USB线连接ZX706且看不到盘符，请点击右下角的『📱 直接扫描 USB 连接的 ZX706 设备』按钮。")
            return

        self.is_running = True
        threading.Thread(target=self._run_folder_scan_thread, args=(folder,), daemon=True).start()

    def _run_folder_scan_thread(self, folder):
        try:
            self.status_var.set("正在遍历文件夹查找音频文件...")
            self.progress.config(mode="indeterminate")
            self.progress.start(10)
            self.log(f"🔍 开始扫描文件夹: {folder}")

            audio_files = []
            for root_dir, _, files in os.walk(folder):
                for f in files:
                    ext = os.path.splitext(f)[1].lower()
                    if ext in AUDIO_EXTENSIONS and not f.startswith("._"):
                        full_path = os.path.join(root_dir, f)
                        rel_dir = os.path.relpath(root_dir, folder)
                        audio_files.append((full_path, f, rel_dir))

            self.progress.stop()
            total = len(audio_files)
            if total == 0:
                self.status_var.set("未找到任何音乐文件。")
                self.log("❌ 在所选文件夹中未找到支持的音频文件。")
                messagebox.showinfo("结果", "在该文件夹下未找到任何音乐文件。")
                self.is_running = False
                return

            self.log(f"✅ 共发现 {total} 首音乐文件，正在解析歌曲信息...")
            self.progress.config(mode="determinate", maximum=total, value=0)

            song_records = []
            for idx, (full_path, filename, rel_dir) in enumerate(audio_files, 1):
                info = self.extract_song_metadata(full_path, filename, rel_dir)
                song_records.append(info)
                if idx % 20 == 0 or idx == total:
                    self.progress["value"] = idx
                    self.status_var.set(f"正在解析歌曲 ({idx}/{total})...")

            self._export_records(song_records)

        except Exception as e:
            self.log(f"❌ 发生错误: {e}")
            messagebox.showerror("错误", f"扫描过程中出现错误：\n{e}")
        finally:
            self.progress.stop()
            self.is_running = False

    def start_mtp_scan(self):
        if self.is_running:
            return
        if not sys.platform.startswith("win"):
            messagebox.showinfo("提示", "MTP 便携设备直连扫描仅适用于 Windows 系统。Mac 用户请安装 Android File Transfer 或使用读卡器选择文件夹。")
            return

        self.is_running = True
        threading.Thread(target=self._run_mtp_scan_thread, daemon=True).start()

    def _run_mtp_scan_thread(self):
        """通过 Windows Shell.Application 接口直接遍历 MTP 便携设备（如 NW-ZX706）"""
        try:
            self.status_var.set("正在连接 USB MTP 设备并扫描音乐文件（700首歌约需10~30秒）...")
            self.progress.config(mode="indeterminate")
            self.progress.start(10)
            self.log("🔍 正在检索『此电脑』中的便携式播放器设备 (如 NW-ZX706 / Walkman)...")

            # 内嵌 PowerShell 脚本，利用系统 Shell.Application 读取 MTP 虚拟目录
            ps_script = r"""
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            $shell = New-Object -ComObject Shell.Application
            $computer = $shell.NameSpace(17)
            $audioExts = @('.mp3','.flac','.wav','.aac','.m4a','.dsf','.dff','.ape','.ogg','.wma','.alac','.aiff','.aif','.opus')
            $results = New-Object System.Collections.Generic.List[Object]

            function Walk-MtpFolder($folder, $pathPrefix) {
                foreach ($item in $folder.Items()) {
                    if ($item.IsFolder) {
                        $subName = $item.Name
                        # 跳过 Android 系统缓存目录以大幅提升扫描速度
                        if ($subName -notin @('Android', '.thumbnails', 'LOST.DIR', 'System Volume Information')) {
                            $subFolder = $item.GetFolder()
                            if ($null -ne $subFolder) {
                                Walk-MtpFolder $subFolder ($pathPrefix + "\" + $subName)
                            }
                        }
                    } else {
                        $name = $item.Name
                        $ext = [System.IO.Path]::GetExtension($name).ToLower()
                        # 兼容部分隐藏扩展名的系统设置，同时检测完整名称
                        if ($audioExts -contains $ext) {
                            $results.Add(@{ name = $name; folder = $pathPrefix })
                        }
                    }
                }
            }

            $devicesFound = @()
            foreach ($dev in $computer.Items()) {
                # 筛选非普通固定盘符的便携设备或名称包含 ZX706/Walkman 的设备
                if ($dev.Path -notmatch '^[A-Z]:\\$' -and $dev.IsFolder) {
                    $devFolder = $dev.GetFolder()
                    if ($null -ne $devFolder) {
                        $devicesFound += $dev.Name
                        Walk-MtpFolder $devFolder $dev.Name
                    }
                }
            }

            $output = @{
                devices = $devicesFound
                files = $results
            }
            $output | ConvertTo-Json -Depth 4 -Compress
            """

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

            proc = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                startupinfo=startupinfo
            )

            self.progress.stop()
            raw_json = proc.stdout.strip()
            if not raw_json:
                self.log("❌ 未能从 USB MTP 设备读取到数据。请确认播放器已解锁屏幕并开启『文件传输 (MTP)』模式。")
                messagebox.showwarning(
                    "未检测到设备或文件",
                    "未能读取到 USB 便携设备中的文件。\n\n请检查：\n"
                    "1. 索尼 ZX706 是否已通过 USB 连接电脑\n"
                    "2. 播放器屏幕上是否已下拉通知栏，将 USB 用途切换为『文件传输』\n"
                    "3. 在『此电脑』中能否看到 NW-ZX706 图标"
                )
                self.is_running = False
                return

            data = json.loads(raw_json)
            devices = data.get("devices", [])
            files = data.get("files", [])
            if isinstance(files, dict):
                files = [files]

            self.log(f"📱 已扫描设备: {', '.join(devices) if devices else '未命名设备'}")
            total = len(files)
            if total == 0:
                self.log("❌ 未在便携设备中发现音频文件。")
                messagebox.showinfo("提示", "已连接设备，但未扫描到支持的音乐文件。请确认播放器已开启『文件传输』模式。")
                self.is_running = False
                return

            self.log(f"✅ 共发现 {total} 首音乐文件，正在整理歌名与歌手...")
            self.progress.config(mode="determinate", maximum=total, value=0)

            song_records = []
            for idx, item in enumerate(files, 1):
                fname = item.get("name", "")
                ffolder = item.get("folder", "")
                info = self.extract_song_metadata(None, fname, ffolder)
                song_records.append(info)
                if idx % 20 == 0 or idx == total:
                    self.progress["value"] = idx

            self._export_records(song_records)

        except Exception as e:
            self.log(f"❌ MTP 扫描出错: {e}")
            messagebox.showerror("错误", f"扫描 USB 设备时出错：\n{e}")
        finally:
            self.progress.stop()
            self.is_running = False

    def _export_records(self, records):
        out_dir = self.output_dir_var.get().strip()
        os.makedirs(out_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # 1. 导出纯文本文件 (.txt) —— 严格按照 "歌曲名字 - 歌手" 每行一条
        txt_path = os.path.join(out_dir, f"ZX706_歌曲名字-歌手_{timestamp}.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            for r in records:
                f.write(f"{r['combined']}\n")

        # 2. 导出 Excel 文件 (.xlsx)
        if HAS_OPENPYXL:
            excel_path = os.path.join(out_dir, f"ZX706_歌单备份_{timestamp}.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.title = "ZX706歌曲清单"

            headers = ["序号", "歌曲名字 - 歌手", "歌曲名称", "歌手", "专辑", "格式", "原文件名", "所在文件夹"]
            ws.append(headers)

            # 美化表头
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for idx, r in enumerate(records, 1):
                ws.append([
                    idx,
                    r["combined"],
                    r["title"],
                    r["artist"],
                    r["album"],
                    r["ext"],
                    r["filename"],
                    r["folder"]
                ])

            # 调整列宽
            col_widths = [8, 38, 25, 20, 22, 10, 35, 30]
            for i, width in enumerate(col_widths, 1):
                col_letter = chr(64 + i)
                ws.column_dimensions[col_letter].width = width

            wb.save(excel_path)
            self.log(f"🎉 导出成功！共备份 {len(records)} 首歌曲：")
            self.log(f"   📊 Excel 表格: {excel_path}")
            self.log(f"   📄 文本清单: {txt_path}\n")
        else:
            # 降级导出带 BOM 的 CSV（可直接用 Excel 双击打开不乱码）
            excel_path = os.path.join(out_dir, f"ZX706_歌单备份_{timestamp}.csv")
            import csv
            with open(excel_path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["序号", "歌曲名字 - 歌手", "歌曲名称", "歌手", "专辑", "格式", "原文件名", "所在文件夹"])
                for idx, r in enumerate(records, 1):
                    writer.writerow([
                        idx, r["combined"], r["title"], r["artist"],
                        r["album"], r["ext"], r["filename"], r["folder"]
                    ])
            self.log(f"🎉 导出成功（CSV格式，可用Excel直接打开）：{excel_path}")

        self.status_var.set(f"导出完成！共导出 {len(records)} 首歌曲。")
        if messagebox.askyesno("导出成功", f"成功导出 {len(records)} 首歌曲！\n\n是否立即打开生成的 Excel 备份文件夹？"):
            self.open_output_dir()


if __name__ == "__main__":
    root = tk.Tk()
    app = ZX706MusicExporterApp(root)
    root.mainloop()
