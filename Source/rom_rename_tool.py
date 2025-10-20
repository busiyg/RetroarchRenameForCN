"""
ROM Batch Renamer (GUI) - 优化精简版
依赖: pip install rapidfuzz pandas
"""
import os
import json
import threading
import re
import webbrowser
from datetime import datetime
from pathlib import Path
from tkinter import Tk, Label, Entry, Button, StringVar, IntVar, filedialog, DISABLED, NORMAL, END, Frame
from tkinter.scrolledtext import ScrolledText
import pandas as pd
from rapidfuzz import process, fuzz

APP_TITLE = "ROM批量重命名工具 (动态CSV匹配版)"

# 扩展名到CSV文件的映射
EXT_TO_CSV = {
    '.gba': 'Nintendo - Game Boy Advance.csv',
    '.gbc': 'Nintendo - Game Boy Color.csv',
    '.gb': 'Nintendo - Game Boy.csv',
    ('.3ds', '.cia'): 'Nintendo - New Nintendo 3DS.csv',
    '.nds': 'Nintendo - Nintendo DS.csv',
    ('.n64', '.z64', '.v64'): 'Nintendo - Nintendo 64.csv',
    '.nes': 'Nintendo - Nintendo Entertainment System.csv',
    ('.sfc', '.smc'): 'Nintendo - Super Nintendo Entertainment System.csv',
    ('.wud', '.wux'): 'Nintendo - Wii U.csv',
    ('.wbfs', '.iso'): 'Nintendo - Wii.csv',
    '.cso': 'Sony - PlayStation Portable.csv',
    ('.bin', '.cue'): 'Sony - PlayStation.csv',
}

class CSVMapper:
    """CSV映射和缓存管理"""
    def __init__(self, csv_root='rom-name-cn-master'):
        self.csv_dir = Path(__file__).parent / csv_root
        self.cache = {}
        self._build_ext_map()
    
    def _build_ext_map(self):
        """展开扩展名映射"""
        self.ext_map = {}
        for key, csv_name in EXT_TO_CSV.items():
            if isinstance(key, tuple):
                for ext in key:
                    self.ext_map[ext.lower()] = csv_name
            else:
                self.ext_map[key.lower()] = csv_name
    
    def get_csv_path(self, ext):
        """根据扩展名获取CSV路径"""
        csv_name = self.ext_map.get(ext.lower())
        if csv_name:
            csv_path = self.csv_dir / csv_name
            return csv_path if csv_path.exists() else None
        return None
    
    def load_mapping(self, csv_path):
        """加载CSV映射(带缓存)"""
        if csv_path not in self.cache:
            try:
                df = pd.read_csv(csv_path, header=None, dtype=str, encoding='utf-8-sig')
                if df.shape[1] < 2:
                    raise ValueError("CSV必须至少包含两列")
                df = df.fillna("")
                eng, cn = df.iloc[:, 0].tolist(), df.iloc[:, 1].tolist()
                self.cache[csv_path] = {
                    'cn_to_eng': {c: e for e, c in zip(eng, cn) if c},
                    'eng_to_cn': {e: c for e, c in zip(eng, cn) if c},
                    'cn_list': [c for c in cn if c]
                }
            except Exception as e:
                raise ValueError(f"读取CSV失败: {e}")
        return self.cache[csv_path]

class FileNameCleaner:
    """文件名清理工具"""
    @staticmethod
    def clean_prefix(name):
        """清理中文前的前缀"""
        match = re.search(r'[\u4e00-\u9fff]', name)
        if not match:
            return name
        pos = match.start()
        prefix = name[:pos].rstrip()
        # 检查是否以单个字母结尾
        if prefix and prefix[-1].isalpha() and (len(prefix) == 1 or not prefix[-2].isalpha()):
            return name[pos:]
        # 如果前缀没有字母则全部移除
        if not any(c.isalpha() for c in prefix):
            return name[pos:]
        return name
    
    @staticmethod
    def clean(name):
        """完整清理流程"""
        name = FileNameCleaner.clean_prefix(name)
        name = re.sub(r'[\[\(].*?[\]\)]', '', name)  # 移除括号内容
        name = re.sub(r'[_\-\+]+', '', name)  # 移除分隔符
        name = re.sub(r'\s+', '', name)  # 移除所有空格
        name = name.strip()
        return re.sub(r'Advance', 'A', name, flags=re.IGNORECASE)

class SmartMatcher:
    """智能匹配器"""
    @staticmethod
    def match(query, choices, threshold):
        """多策略智能匹配"""
        if not query or not choices:
            return None, 0
        
        candidates = process.extract(query, choices, scorer=fuzz.token_set_ratio, limit=5)
        best_match, best_score = None, 0
        
        for candidate, token_score, *_ in candidates:
            # 综合多种匹配策略
            scores = [
                fuzz.token_set_ratio(query, candidate),
                fuzz.ratio(query, candidate),
                fuzz.partial_ratio(query, candidate),
                fuzz.token_sort_ratio(query, candidate)
            ]
            
            # 长度惩罚
            len_ratio = min(len(query), len(candidate)) / max(len(query), len(candidate))
            composite = sum(s * w for s, w in zip(scores, [0.35, 0.25, 0.25, 0.15])) * (len_ratio ** 0.5)
            
            # 子串惩罚
            if len(candidate) < len(query) and candidate in query:
                composite *= 0.7
            
            if composite > best_score:
                best_score, best_match = composite, candidate
        
        return (best_match, best_score) if best_score >= threshold else (None, best_score)

class RenamerApp:
    def __init__(self, master):
        self.master = master
        master.title(APP_TITLE)
        master.resizable(False, False)
        
        self.folder_var = StringVar()
        self.lpl_var = StringVar()
        self.threshold_var = IntVar(value=40)
        self.mapper = CSVMapper()
        self.running = False
        
        self._build_ui()
    
    def _build_ui(self):
        """构建UI"""
        # 文件夹选择
        Label(self.master, text="ROM 文件夹：").grid(row=0, column=0, sticky='w', padx=6, pady=6)
        Entry(self.master, textvariable=self.folder_var, width=60).grid(row=0, column=1, padx=6, pady=6, columnspan=2)
        Button(self.master, text="浏览", command=lambda: self._browse(self.folder_var, True)).grid(row=0, column=3, padx=6)
        
        # LPL选择
        Label(self.master, text="LPL 播放列表：").grid(row=1, column=0, sticky='w', padx=6, pady=6)
        Entry(self.master, textvariable=self.lpl_var, width=60).grid(row=1, column=1, padx=6, pady=6, columnspan=2)
        Button(self.master, text="浏览", command=lambda: self._browse(self.lpl_var, False)).grid(row=1, column=3, padx=6)
        
        # 阈值设置
        Label(self.master, text="匹配阈值 (0-100)：").grid(row=2, column=0, sticky='w', padx=6, pady=6)
        Entry(self.master, textvariable=self.threshold_var, width=8).grid(row=2, column=1, sticky='w', padx=6, pady=6)
        Label(self.master, text="(自动清理文件名前缀并匹配CSV)", fg="gray").grid(row=2, column=1, columnspan=2, sticky='e', padx=6)
        
        # 操作按钮
        self.preview_btn = Button(self.master, text="预览重命名效果", command=self._start_preview, width=18)
        self.preview_btn.grid(row=3, column=0, padx=6, pady=6)
        self.run_btn = Button(self.master, text="执行重命名", command=self._start_rename, width=15)
        self.run_btn.grid(row=3, column=1, sticky='w', padx=6, pady=6)
        self.lpl_btn = Button(self.master, text="转换LPL标签", command=self._start_lpl, width=15)
        self.lpl_btn.grid(row=3, column=2, sticky='w', padx=6, pady=6)
        Button(self.master, text="清空日志", command=self._clear_log, width=10).grid(row=3, column=3, padx=6, pady=6)
        
        # 日志区域
        Label(self.master, text="日志/进度：").grid(row=4, column=0, sticky='nw', padx=6, pady=6)
        self.log = ScrolledText(self.master, width=90, height=20, state=DISABLED)
        self.log.grid(row=4, column=1, columnspan=3, padx=6, pady=6)
        
        # 作者信息
        author_frame = Frame(self.master)
        author_frame.grid(row=5, column=1, columnspan=3, pady=6)
        Label(author_frame, text="作者：").pack(side='left')
        Button(author_frame, text="奇个旦", fg="blue", cursor="hand2", relief="flat",
               command=lambda: webbrowser.open("https://space.bilibili.com/332938511")).pack(side='left')
    
    def _browse(self, var, is_folder):
        """浏览文件/文件夹"""
        result = filedialog.askdirectory() if is_folder else filedialog.askopenfilename(filetypes=[("LPL files", "*.lpl")])
        if result:
            var.set(result)
    
    def _clear_log(self):
        """清空日志"""
        self.log.configure(state=NORMAL)
        self.log.delete(1.0, END)
        self.log.configure(state=DISABLED)
    
    def _log(self, text):
        """写入日志"""
        self.log.configure(state=NORMAL)
        self.log.insert(END, f"[{datetime.now():%H:%M:%S}] {text}\n")
        self.log.see(END)
        self.log.configure(state=DISABLED)
    
    def _validate_and_start(self, callback, *args):
        """验证输入并启动任务"""
        if self.running:
            return
        try:
            threshold = int(self.threshold_var.get())
            if not 0 <= threshold <= 100:
                raise ValueError
        except:
            self._log("错误：阈值需为0-100的整数")
            return
        
        self.run_btn.configure(state=DISABLED)
        self.lpl_btn.configure(state=DISABLED)
        self.preview_btn.configure(state=DISABLED)
        self.running = True
        self.mapper.cache.clear()
        threading.Thread(target=callback, args=(*args, threshold), daemon=True).start()
    
    def _start_preview(self):
        """启动预览模式"""
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            self._log("错误：请选择有效的ROM文件夹")
            return
        self._validate_and_start(self._preview_roms, folder)
    
    def _start_rename(self):
        """启动ROM重命名"""
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            self._log("错误：请选择有效的ROM文件夹")
            return
        self._validate_and_start(self._rename_roms, folder)
    
    def _start_lpl(self):
        """启动LPL转换"""
        lpl_path = self.lpl_var.get().strip()
        if not lpl_path or not os.path.exists(lpl_path):
            self._log("错误：请选择有效的LPL文件")
            return
        self._validate_and_start(self._convert_lpl, lpl_path)
    
    def _preview_roms(self, folder, threshold):
        """预览重命名效果（不实际修改）"""
        from time import time
        start = time()
        self._log("=" * 70)
        self._log("【预览模式】开始分析ROM文件...")
        
        stats = {'total': 0, 'will_rename': 0, 'skipped': 0, 'english': 0, 'no_csv': 0, 'errors': 0}
        platform_stats = {}
        
        for filename in os.listdir(folder):
            if not os.path.isfile(os.path.join(folder, filename)):
                continue
            
            stats['total'] += 1
            name, ext = os.path.splitext(filename)
            
            try:
                # 跳过英文文件
                if not re.search(r'[\u4e00-\u9fff]', name):
                    stats['english'] += 1
                    continue
                
                # 获取CSV映射
                csv_path = self.mapper.get_csv_path(ext)
                if not csv_path:
                    stats['no_csv'] += 1
                    self._log(f"⚠ 未找到CSV: {filename} ({ext})")
                    continue
                
                csv_name = csv_path.name
                platform_stats[csv_name] = platform_stats.get(csv_name, 0) + 1
                
                # 匹配预览
                mapping = self.mapper.load_mapping(csv_path)
                cleaned = FileNameCleaner.clean(name)
                match, score = SmartMatcher.match(cleaned, mapping['cn_list'], threshold)
                
                if match and (eng := mapping['cn_to_eng'].get(match)):
                    new_name = self._unique_path(folder, eng + ext)
                    stats['will_rename'] += 1
                    self._log(f"✓ [预览] {filename}\n  → {new_name}\n  [CSV: {csv_name}, 分数: {score:.1f}]")
                else:
                    stats['skipped'] += 1
                    self._log(f"✗ 将跳过: {filename} (分数:{score:.1f})")
            
            except Exception as e:
                stats['errors'] += 1
                self._log(f"✗ 错误: {filename} - {e}")
        
        # 输出统计
        self._log("=" * 70)
        self._log(f"预览完成! 耗时: {time()-start:.1f}s")
        self._log(f"总计: {stats['total']} | 将重命名: {stats['will_rename']} | 跳过英文: {stats['english']}")
        self._log(f"将跳过: {stats['skipped']} | 无CSV: {stats['no_csv']} | 错误: {stats['errors']}")
        
        if platform_stats:
            self._log("\n使用的CSV统计:")
            for csv, count in sorted(platform_stats.items()):
                self._log(f"  • {csv}: {count} 个文件")
        
        self._log("\n⚡ 提示：如果预览效果满意，点击「执行重命名」按钮正式重命名文件")
        self._finish()
    
    def _rename_roms(self, folder, threshold):
        """重命名ROM文件"""
        from time import time
        start = time()
        self._log("=" * 70)
        self._log("开始重命名ROM文件...")
        
        stats = {'total': 0, 'renamed': 0, 'skipped': 0, 'english': 0, 'no_csv': 0, 'errors': 0}
        platform_stats = {}
        
        for filename in os.listdir(folder):
            if not os.path.isfile(os.path.join(folder, filename)):
                continue
            
            stats['total'] += 1
            name, ext = os.path.splitext(filename)
            
            try:
                # 跳过英文文件
                if not re.search(r'[\u4e00-\u9fff]', name):
                    stats['english'] += 1
                    continue
                
                # 获取CSV映射
                csv_path = self.mapper.get_csv_path(ext)
                if not csv_path:
                    stats['no_csv'] += 1
                    self._log(f"⚠ 未找到CSV: {filename} ({ext})")
                    continue
                
                csv_name = csv_path.name
                platform_stats[csv_name] = platform_stats.get(csv_name, 0) + 1
                
                # 匹配并重命名
                mapping = self.mapper.load_mapping(csv_path)
                cleaned = FileNameCleaner.clean(name)
                match, score = SmartMatcher.match(cleaned, mapping['cn_list'], threshold)
                
                if match and (eng := mapping['cn_to_eng'].get(match)):
                    new_name = self._unique_path(folder, eng + ext)
                    os.rename(os.path.join(folder, filename), os.path.join(folder, new_name))
                    stats['renamed'] += 1
                    self._log(f"✓ {filename}\n  → {new_name}\n  [CSV: {csv_name}, 分数: {score:.1f}]")
                else:
                    stats['skipped'] += 1
                    self._log(f"✗ 跳过: {filename} (分数:{score:.1f})")
            
            except Exception as e:
                stats['errors'] += 1
                self._log(f"✗ 错误: {filename} - {e}")
        
        # 输出统计
        self._log("=" * 70)
        self._log(f"完成! 耗时: {time()-start:.1f}s")
        self._log(f"总计: {stats['total']} | 成功: {stats['renamed']} | 跳过英文: {stats['english']}")
        self._log(f"未匹配: {stats['skipped']} | 无CSV: {stats['no_csv']} | 错误: {stats['errors']}")
        
        if platform_stats:
            self._log("\n使用的CSV统计:")
            for csv, count in sorted(platform_stats.items()):
                self._log(f"  • {csv}: {count} 个文件")
        
        self._finish()
    
    def _convert_lpl(self, lpl_path, threshold):
        """转换LPL播放列表"""
        from time import time
        start = time()
        self._log("=" * 70)
        self._log(f"开始转换LPL: {Path(lpl_path).name}")
        
        try:
            with open(lpl_path, 'r', encoding='utf-8') as f:
                lpl = json.load(f)
            
            if 'items' not in lpl:
                self._log("✗ 错误：LPL格式不正确")
                self._finish()
                return
            
            stats = {'total': len(lpl['items']), 'converted': 0, 'skipped': 0, 'no_csv': 0}
            platform_stats = {}
            
            for item in lpl['items']:
                label = item.get('label', '')
                ext = Path(item.get('path', '')).suffix
                
                csv_path = self.mapper.get_csv_path(ext)
                if not csv_path:
                    stats['no_csv'] += 1
                    self._log(f"⚠ 未找到CSV: {label} ({ext})")
                    continue
                
                csv_name = csv_path.name
                platform_stats[csv_name] = platform_stats.get(csv_name, 0) + 1
                
                # 匹配中文名
                mapping = self.mapper.load_mapping(csv_path)
                cleaned = FileNameCleaner.clean(label)
                best_match, best_score = None, 0
                
                for eng, cn in mapping['eng_to_cn'].items():
                    if cn:
                        score = fuzz.token_set_ratio(cleaned, FileNameCleaner.clean(eng))
                        if score > best_score:
                            best_score, best_match = score, cn
                
                if best_match and best_score >= threshold:
                    item['label'] = best_match
                    stats['converted'] += 1
                    self._log(f"✓ {label}\n  → {best_match} [{csv_name}, {best_score:.1f}]")
                else:
                    stats['skipped'] += 1
                    self._log(f"⊙ 保持: {label} ({best_score:.1f})")
            
            # 保存到桌面
            desktop = Path.home() / "Desktop"
            clean_name = Path(lpl_path).name.replace('_', ' ')
            clean_name = re.sub(r'\[.*?\]', '', clean_name).strip()
            save_path = desktop / clean_name
            
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(lpl, f, ensure_ascii=False, indent=2)
            
            self._log("=" * 70)
            self._log(f"完成! 耗时: {time()-start:.1f}s")
            self._log(f"总计: {stats['total']} | 已转换: {stats['converted']} | 保持: {stats['skipped']} | 无CSV: {stats['no_csv']}")
            self._log(f"已保存到桌面: {clean_name}")
            
            if platform_stats:
                self._log("\n使用的CSV统计:")
                for csv, count in sorted(platform_stats.items()):
                    self._log(f"  • {csv}: {count} 个条目")
        
        except Exception as e:
            self._log(f"✗ 处理失败: {e}")
        
        self._finish()
    
    def _unique_path(self, folder, filename):
        """生成唯一文件名"""
        base, ext = os.path.splitext(filename)
        candidate, i = filename, 1
        while os.path.exists(os.path.join(folder, candidate)):
            candidate = f"{base} ({i}){ext}"
            i += 1
        return candidate
    
    def _finish(self):
        """完成任务"""
        self.running = False
        self.run_btn.configure(state=NORMAL)
        self.lpl_btn.configure(state=NORMAL)
        self.preview_btn.configure(state=NORMAL)

if __name__ == '__main__':
    root = Tk()
    app = RenamerApp(root)
    root.mainloop()