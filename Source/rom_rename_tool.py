"""
ROM Batch Renamer (GUI) - 动态CSV匹配版本
主要改进：
- 每个文件根据自身扩展名动态匹配对应的CSV映射表
- 支持混合平台ROM文件夹
- LPL文件根据每个条目的文件扩展名单独匹配CSV
- 自动在rom-name-cn-master目录查找对应CSV
依赖：
    pip install rapidfuzz pandas
"""
import os
import json
import threading
import time
import traceback
import re
import webbrowser
from datetime import datetime
from tkinter import Tk, Label, Entry, Button, StringVar, IntVar, filedialog, DISABLED, NORMAL, END
from tkinter.scrolledtext import ScrolledText
import pandas as pd
from rapidfuzz import process, fuzz

APP_TITLE = "ROM批量重命名工具 (动态CSV匹配版)"

# 扩展名到CSV文件的映射规则
EXTENSION_TO_CSV = {
    # Game Boy Advance
    '.gba': 'Nintendo - Game Boy Advance.csv',
    
    # Game Boy Color
    '.gbc': 'Nintendo - Game Boy Color.csv',
    
    # Game Boy
    '.gb': 'Nintendo - Game Boy.csv',
    
    # Nintendo 3DS
    '.3ds': 'Nintendo - New Nintendo 3DS.csv',
    '.cia': 'Nintendo - New Nintendo 3DS.csv',
    
    # Nintendo DS
    '.nds': 'Nintendo - Nintendo DS.csv',
    
    # Nintendo 64
    '.n64': 'Nintendo - Nintendo 64.csv',
    '.z64': 'Nintendo - Nintendo 64.csv',
    '.v64': 'Nintendo - Nintendo 64.csv',
    
    # NES
    '.nes': 'Nintendo - Nintendo Entertainment System.csv',
    
    # SNES
    '.sfc': 'Nintendo - Super Nintendo Entertainment System.csv',
    '.smc': 'Nintendo - Super Nintendo Entertainment System.csv',
    
    # Wii U
    '.wud': 'Nintendo - Wii U.csv',
    '.wux': 'Nintendo - Wii U.csv',
    
    # Wii
    '.wbfs': 'Nintendo - Wii.csv',
    '.iso': 'Nintendo - Wii.csv',
    
    # PlayStation Portable
    '.cso': 'Sony - PlayStation Portable.csv',
    
    # PlayStation
    '.bin': 'Sony - PlayStation.csv',
    '.cue': 'Sony - PlayStation.csv',
}

def get_csv_path_for_extension(ext, csv_root='rom-name-cn-master'):
    """
    根据文件扩展名获取对应的CSV文件路径
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_dir = os.path.join(script_dir, csv_root)
    
    if not os.path.exists(csv_dir):
        return None
    
    ext = ext.lower()
    if ext in EXTENSION_TO_CSV:
        csv_filename = EXTENSION_TO_CSV[ext]
        csv_path = os.path.join(csv_dir, csv_filename)
        if os.path.exists(csv_path):
            return csv_path
    
    return None

def read_mapping(csv_path):
    """读取CSV映射表，返回中文→英文和英文→中文两个字典"""
    try:
        df = pd.read_csv(csv_path, header=None, dtype=str)
    except Exception:
        df = pd.read_csv(csv_path, header=None, dtype=str, encoding='utf-8-sig')
    if df.shape[1] < 2:
        raise ValueError("CSV 必须至少包含两列：英文名, 中文名")
    df = df.fillna("")
    eng = df.iloc[:, 0].astype(str).tolist()
    cn = df.iloc[:, 1].astype(str).tolist()
    
    cn_to_eng = {cn_name: eng_name for eng_name, cn_name in zip(eng, cn) if cn_name}
    eng_to_cn = {eng_name: cn_name for eng_name, cn_name in zip(eng, cn) if cn_name}
    cn_list = [cn_name for cn_name in cn if cn_name]
    
    return cn_to_eng, eng_to_cn, cn_list

def ensure_unique_path(folder, filename):
    base, ext = os.path.splitext(filename)
    candidate = filename
    i = 1
    while os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{base} ({i}){ext}"
        i += 1
    return candidate

def contains_chinese(text):
    """检测文本是否包含中文字符"""
    return bool(re.search(r'[\u4e00-\u9fff]', text))

def clean_filename(name):
    """深度清理文件名"""
    name = re.sub(r'\[.*?\]', '', name)
    name = re.sub(r'\(.*?\)', '', name)
    name = re.sub(r'[_\-\+]+', ' ', name)
    name = re.sub(r'\s+', ' ', name)
    name = name.strip()
    return name

def smart_match(query, choices, threshold):
    """智能匹配算法"""
    if not query or not choices:
        return None, 0
    
    candidates = process.extract(
        query, 
        choices, 
        scorer=fuzz.token_set_ratio,
        limit=5
    )
    
    best_match = None
    best_score = 0
    
    for candidate, token_score, *_ in candidates:
        partial = fuzz.partial_ratio(query, candidate)
        ratio = fuzz.ratio(query, candidate)
        token_sort = fuzz.token_sort_ratio(query, candidate)
        
        len_ratio = min(len(query), len(candidate)) / max(len(query), len(candidate))
        len_penalty = len_ratio ** 0.5
        
        if len(candidate) < len(query) * 0.5:
            partial_weight = 0.2
        else:
            partial_weight = 0.35
        
        composite_score = (
            token_score * 0.35 +
            ratio * 0.25 +
            partial * partial_weight +
            token_sort * 0.15
        ) * len_penalty
        
        if len(candidate) < len(query) and candidate in query:
            composite_score *= 0.7
        
        if composite_score > best_score:
            best_score = composite_score
            best_match = candidate
    
    if best_score >= threshold:
        return best_match, best_score
    return None, best_score

class CSVCache:
    """CSV缓存类，避免重复读取相同的CSV文件"""
    def __init__(self):
        self.cache = {}
    
    def get_mapping(self, csv_path):
        if csv_path not in self.cache:
            self.cache[csv_path] = read_mapping(csv_path)
        return self.cache[csv_path]
    
    def clear(self):
        self.cache.clear()

class RenamerApp:
    def __init__(self, master):
        self.master = master
        master.title(APP_TITLE)
        master.resizable(False, False)
        
        self.folder_var = StringVar()
        self.lpl_var = StringVar()
        self.threshold_var = IntVar(value=40)
        self.csv_cache = CSVCache()
        
        # ROM文件夹选择
        Label(master, text="ROM 文件夹：").grid(row=0, column=0, sticky='w', padx=6, pady=6)
        self.folder_entry = Entry(master, textvariable=self.folder_var, width=60)
        self.folder_entry.grid(row=0, column=1, padx=6, pady=6, columnspan=2)
        Button(master, text="浏览", command=self.browse_folder).grid(row=0, column=3, padx=6)
        
        # LPL播放列表选择
        Label(master, text="LPL 播放列表：").grid(row=1, column=0, sticky='w', padx=6, pady=6)
        self.lpl_entry = Entry(master, textvariable=self.lpl_var, width=60)
        self.lpl_entry.grid(row=1, column=1, padx=6, pady=6, columnspan=2)
        Button(master, text="浏览", command=self.browse_lpl).grid(row=1, column=3, padx=6)
        
        # 匹配阈值
        Label(master, text="匹配阈值 (0-100)：").grid(row=2, column=0, sticky='w', padx=6, pady=6)
        self.threshold_entry = Entry(master, textvariable=self.threshold_var, width=8)
        self.threshold_entry.grid(row=2, column=1, sticky='w', padx=6, pady=6)
        
        Label(master, text="(每个文件根据扩展名自动匹配对应CSV)", fg="gray").grid(
            row=2, column=1, columnspan=2, sticky='e', padx=6
        )
        
        # 按钮区域
        self.run_button = Button(master, text="重命名ROM文件", command=self.start_rename_roms, width=18)
        self.run_button.grid(row=3, column=0, padx=6, pady=6)
        
        self.lpl_button = Button(master, text="转换LPL标签为中文", command=self.start_convert_lpl, width=18)
        self.lpl_button.grid(row=3, column=1, sticky='w', padx=6, pady=6)
        
        self.clear_log_button = Button(master, text="清空日志", command=self.clear_log, width=10)
        self.clear_log_button.grid(row=3, column=2, padx=6, pady=6)
        
        # 日志区域
        Label(master, text="日志/进度：").grid(row=4, column=0, sticky='nw', padx=6, pady=6)
        self.log = ScrolledText(master, width=90, height=20, state=DISABLED)
        self.log.grid(row=4, column=1, columnspan=3, padx=6, pady=6)
        
        # 作者信息
        from tkinter import Frame
        author_frame = Frame(master)
        author_frame.grid(row=5, column=1, columnspan=3, pady=6)
        
        Label(author_frame, text="作者：").pack(side='left')
        author_button = Button(
            author_frame, 
            text="奇个旦", 
            fg="blue", 
            cursor="hand2",
            relief="flat",
            command=lambda: webbrowser.open("https://space.bilibili.com/332938511")
        )
        author_button.pack(side='left')
        
        self.running = False
    
    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)
    
    def browse_lpl(self):
        p = filedialog.askopenfilename(filetypes=[("LPL files", "*.lpl"), ("All files", "*")])
        if p:
            self.lpl_var.set(p)
    
    def clear_log(self):
        self.log.configure(state=NORMAL)
        self.log.delete(1.0, END)
        self.log.configure(state=DISABLED)
    
    def log_write(self, text):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log.configure(state=NORMAL)
        self.log.insert(END, f"[{timestamp}] {text}\n")
        self.log.see(END)
        self.log.configure(state=DISABLED)
    
    def start_rename_roms(self):
        if self.running:
            return
        folder = self.folder_var.get().strip()
        
        try:
            threshold = int(self.threshold_var.get())
        except Exception:
            self.log_write("错误：阈值需为整数（0-100）。")
            return
        if not folder or not os.path.isdir(folder):
            self.log_write("错误：请选择有效的 ROM 文件夹。")
            return
        if threshold < 0 or threshold > 100:
            self.log_write("错误：阈值需在 0 到 100 之间。")
            return
        
        self.run_button.configure(state=DISABLED)
        self.lpl_button.configure(state=DISABLED)
        self.running = True
        self.csv_cache.clear()
        t = threading.Thread(target=self.run_renamer, args=(folder, threshold), daemon=True)
        t.start()
    
    def start_convert_lpl(self):
        if self.running:
            return
        
        lpl_path = self.lpl_var.get().strip()
        
        try:
            threshold = int(self.threshold_var.get())
        except Exception:
            self.log_write("错误：阈值需为整数（0-100）。")
            return
        
        if not lpl_path or not os.path.exists(lpl_path):
            self.log_write("错误：请选择有效的 LPL 文件。")
            return
        
        self.run_button.configure(state=DISABLED)
        self.lpl_button.configure(state=DISABLED)
        self.running = True
        self.csv_cache.clear()
        t = threading.Thread(target=self.run_lpl_converter, args=(lpl_path, threshold), daemon=True)
        t.start()
    
    def run_renamer(self, folder, threshold):
        start_time = time.time()
        self.log_write("=" * 70)
        self.log_write(f"开始重命名ROM文件（每个文件动态匹配CSV）...")
        self.log_write(f"ROM文件夹: {folder}")
        
        total = 0
        renamed = 0
        skipped = 0
        skipped_english = 0
        no_csv = 0
        errors = 0
        
        # 统计各平台文件数量
        platform_stats = {}
        
        files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
        for filename in files:
            total += 1
            src_path = os.path.join(folder, filename)
            name, ext = os.path.splitext(filename)
            
            try:
                # 检测是否包含中文
                if not contains_chinese(name):
                    skipped_english += 1
                    continue
                
                # 根据扩展名获取CSV路径
                csv_path = get_csv_path_for_extension(ext)
                if not csv_path:
                    no_csv += 1
                    self.log_write(f"⚠ 未找到CSV映射: {filename} (扩展名: {ext})")
                    continue
                
                # 统计平台
                csv_name = os.path.basename(csv_path)
                platform_stats[csv_name] = platform_stats.get(csv_name, 0) + 1
                
                # 获取映射（使用缓存）
                cn_to_eng, _, cn_list = self.csv_cache.get_mapping(csv_path)
                
                cleaned_name = clean_filename(name)
                match, score = smart_match(cleaned_name, cn_list, threshold)
                
                if match:
                    eng_name = cn_to_eng.get(match)
                    if not eng_name:
                        self.log_write(f"⚠ 找到匹配但无英文名：{match}")
                        skipped += 1
                        continue
                    
                    new_filename = eng_name + ext
                    new_filename = ensure_unique_path(folder, new_filename)
                    new_path = os.path.join(folder, new_filename)
                    os.rename(src_path, new_path)
                    renamed += 1
                    self.log_write(f"✓ {filename}")
                    self.log_write(f"  → {new_filename}")
                    self.log_write(f"  [CSV: {csv_name}, 匹配: {match}, 分数: {score:.1f}]")
                else:
                    skipped += 1
                    self.log_write(f"✗ 跳过: {filename} (最高分:{score:.1f}, CSV: {csv_name})")
            except Exception as e:
                errors += 1
                self.log_write(f"✗ 错误: {filename} - {e}")
        
        elapsed = time.time() - start_time
        self.log_write("=" * 70)
        self.log_write(f"完成!")
        self.log_write(f"总计: {total} | 成功: {renamed} | 跳过英文: {skipped_english}")
        self.log_write(f"未匹配: {skipped} | 无CSV映射: {no_csv} | 错误: {errors}")
        self.log_write(f"耗时: {elapsed:.1f}s")
        
        if platform_stats:
            self.log_write("\n使用的CSV统计:")
            for csv_name, count in sorted(platform_stats.items()):
                self.log_write(f"  • {csv_name}: {count} 个文件")
        
        self.running = False
        self.run_button.configure(state=NORMAL)
        self.lpl_button.configure(state=NORMAL)
    
    def run_lpl_converter(self, lpl_path, threshold):
        start_time = time.time()
        self.log_write("=" * 70)
        self.log_write(f"开始转换LPL播放列表（每个条目动态匹配CSV）...")
        self.log_write(f"LPL文件: {os.path.basename(lpl_path)}")
        
        try:
            # 读取LPL文件
            with open(lpl_path, 'r', encoding='utf-8') as f:
                lpl_data = json.load(f)
            
            if 'items' not in lpl_data:
                self.log_write("✗ 错误：LPL文件格式不正确，缺少'items'字段")
                self.running = False
                self.run_button.configure(state=NORMAL)
                self.lpl_button.configure(state=NORMAL)
                return
            
            total = len(lpl_data['items'])
            converted = 0
            skipped = 0
            no_csv = 0
            
            # 统计各平台
            platform_stats = {}
            
            # 创建备份
            backup_path = lpl_path + '.backup'
            with open(backup_path, 'w', encoding='utf-8') as f:
                json.dump(lpl_data, f, ensure_ascii=False, indent=2)
            self.log_write(f"✓ 已创建备份: {os.path.basename(backup_path)}")
            
            # 转换每个条目
            for item in lpl_data['items']:
                original_label = item.get('label', '')
                file_path = item.get('path', '')
                
                # 从路径获取扩展名
                _, ext = os.path.splitext(file_path)
                
                # 根据扩展名获取CSV路径
                csv_path = get_csv_path_for_extension(ext)
                if not csv_path:
                    no_csv += 1
                    self.log_write(f"⚠ 未找到CSV映射: {original_label} (扩展名: {ext})")
                    continue
                
                # 统计平台
                csv_name = os.path.basename(csv_path)
                platform_stats[csv_name] = platform_stats.get(csv_name, 0) + 1
                
                # 获取映射（使用缓存）
                _, eng_to_cn, _ = self.csv_cache.get_mapping(csv_path)
                
                # 清理label
                cleaned_label = clean_filename(original_label)
                
                # 在映射中查找匹配
                best_match = None
                best_score = 0
                
                for eng_name, cn_name in eng_to_cn.items():
                    if not cn_name:
                        continue
                    
                    cleaned_eng = clean_filename(eng_name)
                    score = fuzz.token_set_ratio(cleaned_label, cleaned_eng)
                    
                    if score > best_score:
                        best_score = score
                        best_match = cn_name
                
                if best_match and best_score >= 85:
                    item['label'] = best_match
                    converted += 1
                    self.log_write(f"✓ {original_label}")
                    self.log_write(f"  → {best_match}")
                    self.log_write(f"  [CSV: {csv_name}, 分数: {best_score:.1f}]")
                else:
                    skipped += 1
                    self.log_write(f"⊙ 保持原样: {original_label} (最高分:{best_score:.1f}, CSV: {csv_name})")
            
            # 写回LPL文件
            with open(lpl_path, 'w', encoding='utf-8') as f:
                json.dump(lpl_data, f, ensure_ascii=False, indent=2)
            
            elapsed = time.time() - start_time
            self.log_write("=" * 70)
            self.log_write(f"完成!")
            self.log_write(f"总计: {total} | 已转换: {converted} | 保持原样: {skipped} | 无CSV映射: {no_csv}")
            self.log_write(f"耗时: {elapsed:.1f}s")
            self.log_write(f"原文件已更新，备份保存在: {os.path.basename(backup_path)}")
            
            if platform_stats:
                self.log_write("\n使用的CSV统计:")
                for csv_name, count in sorted(platform_stats.items()):
                    self.log_write(f"  • {csv_name}: {count} 个条目")
            
        except json.JSONDecodeError as e:
            self.log_write(f"✗ JSON解析错误: {e}")
        except Exception as e:
            self.log_write(f"✗ 处理失败: {e}")
            self.log_write(traceback.format_exc())
        
        self.running = False
        self.run_button.configure(state=NORMAL)
        self.lpl_button.configure(state=NORMAL)

if __name__ == '__main__':
    root = Tk()
    app = RenamerApp(root)
    root.mainloop()