"""
ROM Batch Renamer (GUI) - 自动匹配CSV版本
主要改进：
- 根据ROM文件扩展名自动匹配对应的CSV映射表
- 支持多种主机平台的扩展名识别
- 自动在rom-name-cn-master目录查找对应CSV
依赖：
    pip install rapidfuzz pandas
"""
import os
import threading
import time
import traceback
import re
from datetime import datetime
from tkinter import Tk, Label, Entry, Button, StringVar, IntVar, filedialog, DISABLED, NORMAL, END
from tkinter.scrolledtext import ScrolledText
import pandas as pd
from rapidfuzz import process, fuzz

APP_TITLE = "ROM Batch Renamer (智能CSV匹配)"

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
    '.iso': 'Nintendo - Wii.csv',  # 注意：ISO可能对应多个平台
    
    # PlayStation Portable
    '.cso': 'Sony - PlayStation Portable.csv',
    
    # PlayStation
    '.bin': 'Sony - PlayStation.csv',
    '.cue': 'Sony - PlayStation.csv',
}

# 支持通过文件夹名称推断平台
FOLDER_NAME_TO_CSV = {
    'gba': 'Nintendo - Game Boy Advance.csv',
    'game boy advance': 'Nintendo - Game Boy Advance.csv',
    'gbc': 'Nintendo - Game Boy Color.csv',
    'game boy color': 'Nintendo - Game Boy Color.csv',
    'gb': 'Nintendo - Game Boy.csv',
    'game boy': 'Nintendo - Game Boy.csv',
    '3ds': 'Nintendo - New Nintendo 3DS.csv',
    'nds': 'Nintendo - Nintendo DS.csv',
    'n64': 'Nintendo - Nintendo 64.csv',
    'nintendo 64': 'Nintendo - Nintendo 64.csv',
    'nes': 'Nintendo - Nintendo Entertainment System.csv',
    'snes': 'Nintendo - Super Nintendo Entertainment System.csv',
    'super nintendo': 'Nintendo - Super Nintendo Entertainment System.csv',
    'wii u': 'Nintendo - Wii U.csv',
    'wiiu': 'Nintendo - Wii U.csv',
    'wii': 'Nintendo - Wii.csv',
    'psp': 'Sony - PlayStation Portable.csv',
    'playstation portable': 'Sony - PlayStation Portable.csv',
    'ps1': 'Sony - PlayStation.csv',
    'psx': 'Sony - PlayStation.csv',
    'playstation': 'Sony - PlayStation.csv',
}

def detect_csv_file(folder_path, csv_root='rom-name-cn-master'):
    """
    自动检测应该使用哪个CSV文件
    检测策略：
    1. 扫描文件夹中的文件扩展名
    2. 统计最常见的扩展名
    3. 根据映射规则返回对应的CSV文件路径
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_dir = os.path.join(script_dir, csv_root)
    
    if not os.path.exists(csv_dir):
        return None, f"CSV目录不存在: {csv_dir}"
    
    # 统计扩展名
    ext_count = {}
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    
    for filename in files:
        _, ext = os.path.splitext(filename)
        ext = ext.lower()
        if ext:
            ext_count[ext] = ext_count.get(ext, 0) + 1
    
    if not ext_count:
        # 尝试通过文件夹名称推断
        folder_name = os.path.basename(folder_path).lower()
        for keyword, csv_file in FOLDER_NAME_TO_CSV.items():
            if keyword in folder_name:
                csv_path = os.path.join(csv_dir, csv_file)
                if os.path.exists(csv_path):
                    return csv_path, f"根据文件夹名称 '{folder_name}' 匹配到: {csv_file}"
        return None, "未找到ROM文件，无法自动检测平台"
    
    # 找到最常见的扩展名
    most_common_ext = max(ext_count, key=ext_count.get)
    
    # 查找对应的CSV
    if most_common_ext in EXTENSION_TO_CSV:
        csv_filename = EXTENSION_TO_CSV[most_common_ext]
        csv_path = os.path.join(csv_dir, csv_filename)
        
        if os.path.exists(csv_path):
            return csv_path, f"检测到扩展名 '{most_common_ext}' ({ext_count[most_common_ext]}个文件)，匹配到: {csv_filename}"
        else:
            return None, f"找到映射规则 {csv_filename}，但文件不存在: {csv_path}"
    
    # 如果没有直接匹配，尝试通过文件夹名称
    folder_name = os.path.basename(folder_path).lower()
    for keyword, csv_file in FOLDER_NAME_TO_CSV.items():
        if keyword in folder_name:
            csv_path = os.path.join(csv_dir, csv_file)
            if os.path.exists(csv_path):
                return csv_path, f"扩展名 '{most_common_ext}' 无映射，但根据文件夹名称匹配到: {csv_file}"
    
    return None, f"扩展名 '{most_common_ext}' 未在映射规则中，且无法通过文件夹名称推断平台"

def read_mapping(csv_path):
    try:
        df = pd.read_csv(csv_path, header=None, dtype=str)
    except Exception:
        df = pd.read_csv(csv_path, header=None, dtype=str, encoding='utf-8-sig')
    if df.shape[1] < 2:
        raise ValueError("CSV 必须至少包含两列：英文名, 中文名")
    df = df.fillna("")
    eng = df.iloc[:, 0].astype(str).tolist()
    cn = df.iloc[:, 1].astype(str).tolist()
    mapping = {cn_name: eng_name for eng_name, cn_name in zip(eng, cn)}
    return mapping, cn, eng

def ensure_unique_path(folder, filename):
    base, ext = os.path.splitext(filename)
    candidate = filename
    i = 1
    while os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{base} ({i}){ext}"
        i += 1
    return candidate

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

class RenamerApp:
    def __init__(self, master):
        self.master = master
        master.title(APP_TITLE)
        master.resizable(False, False)
        
        self.folder_var = StringVar()
        self.csv_var = StringVar()
        self.threshold_var = IntVar(value=70)
        self.auto_detect_var = IntVar(value=1)  # 默认开启自动检测
        
        Label(master, text="ROM 文件夹：").grid(row=0, column=0, sticky='w', padx=6, pady=6)
        self.folder_entry = Entry(master, textvariable=self.folder_var, width=50)
        self.folder_entry.grid(row=0, column=1, padx=6, pady=6)
        Button(master, text="浏览", command=self.browse_folder).grid(row=0, column=2, padx=6)
        
        Label(master, text="CSV 映射表：").grid(row=1, column=0, sticky='w', padx=6, pady=6)
        self.csv_entry = Entry(master, textvariable=self.csv_var, width=50)
        self.csv_entry.grid(row=1, column=1, padx=6, pady=6)
        Button(master, text="浏览", command=self.browse_csv).grid(row=1, column=2, padx=6)
        
        from tkinter import Checkbutton
        self.auto_detect_check = Checkbutton(
            master, 
            text="自动检测CSV（根据ROM扩展名）", 
            variable=self.auto_detect_var,
            command=self.toggle_auto_detect
        )
        self.auto_detect_check.grid(row=1, column=3, sticky='w', padx=6)
        
        Label(master, text="匹配阈值 (0-100)：").grid(row=2, column=0, sticky='w', padx=6, pady=6)
        self.threshold_entry = Entry(master, textvariable=self.threshold_var, width=8)
        self.threshold_entry.grid(row=2, column=1, sticky='w', padx=6, pady=6)
        
        self.run_button = Button(master, text="开始重命名", command=self.start)
        self.run_button.grid(row=3, column=0, padx=6, pady=6)
        
        self.clear_log_button = Button(master, text="清空日志", command=self.clear_log)
        self.clear_log_button.grid(row=3, column=1, sticky='w', padx=6, pady=6)
        
        Label(master, text="日志/进度：").grid(row=4, column=0, sticky='nw', padx=6, pady=6)
        self.log = ScrolledText(master, width=82, height=18, state=DISABLED)
        self.log.grid(row=4, column=1, columnspan=3, padx=6, pady=6)
        
        self.running = False
        self.toggle_auto_detect()  # 初始化状态
    
    def toggle_auto_detect(self):
        """切换自动检测模式"""
        if self.auto_detect_var.get():
            self.csv_entry.configure(state=DISABLED)
            self.csv_var.set("(将自动检测)")
        else:
            self.csv_entry.configure(state=NORMAL)
            if self.csv_var.get() == "(将自动检测)":
                self.csv_var.set("")
    
    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)
            # 如果开启自动检测，立即尝试检测CSV
            if self.auto_detect_var.get():
                self.try_auto_detect(folder)
    
    def try_auto_detect(self, folder):
        """尝试自动检测CSV"""
        csv_path, message = detect_csv_file(folder)
        if csv_path:
            self.csv_var.set(csv_path)
            self.log_write(f"✓ {message}")
        else:
            self.csv_var.set("(检测失败)")
            self.log_write(f"✗ {message}")
    
    def browse_csv(self):
        p = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*")])
        if p:
            self.csv_var.set(p)
    
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
    
    def start(self):
        if self.running:
            return
        folder = self.folder_var.get().strip()
        csv_path = self.csv_var.get().strip()
        
        # 处理自动检测
        if self.auto_detect_var.get():
            if csv_path == "(将自动检测)" or csv_path == "(检测失败)" or not csv_path:
                detected_csv, message = detect_csv_file(folder)
                if detected_csv:
                    csv_path = detected_csv
                    self.csv_var.set(csv_path)
                    self.log_write(f"✓ {message}")
                else:
                    self.log_write(f"✗ 自动检测失败: {message}")
                    return
        
        try:
            threshold = int(self.threshold_var.get())
        except Exception:
            self.log_write("错误：阈值需为整数（0-100）。")
            return
        if not folder or not os.path.isdir(folder):
            self.log_write("错误：请选择有效的 ROM 文件夹。")
            return
        if not csv_path or not os.path.exists(csv_path):
            self.log_write("错误：请选择有效的 CSV 文件。")
            return
        if threshold < 0 or threshold > 100:
            self.log_write("错误：阈值需在 0 到 100 之间。")
            return
        
        self.run_button.configure(state=DISABLED)
        self.running = True
        t = threading.Thread(target=self.run_renamer, args=(folder, csv_path, threshold), daemon=True)
        t.start()
    
    def run_renamer(self, folder, csv_path, threshold):
        start_time = time.time()
        self.log_write("=" * 60)
        self.log_write(f"开始重命名...")
        self.log_write(f"使用CSV: {os.path.basename(csv_path)}")
        try:
            mapping, cn_list, *_ = read_mapping(csv_path)
        except Exception as e:
            self.log_write(f"读取 CSV 失败：{e}")
            self.running = False
            self.run_button.configure(state=NORMAL)
            return
        
        total = 0
        renamed = 0
        skipped = 0
        errors = 0
        
        files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
        for filename in files:
            total += 1
            src_path = os.path.join(folder, filename)
            name, ext = os.path.splitext(filename)
            
            try:
                cleaned_name = clean_filename(name)
                match, score = smart_match(cleaned_name, cn_list, threshold)
                
                if match:
                    eng_name = mapping.get(match)
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
                    self.log_write(f"  → {new_filename} (匹配:{match}, 分数:{score:.1f})")
                else:
                    skipped += 1
                    self.log_write(f"✗ 跳过: {filename} (最高分:{score:.1f})")
            except Exception as e:
                errors += 1
                self.log_write(f"✗ 错误: {filename} - {e}")
        
        elapsed = time.time() - start_time
        self.log_write("=" * 60)
        self.log_write(f"完成! 总计:{total} | 成功:{renamed} | 跳过:{skipped} | 错误:{errors} | 耗时:{elapsed:.1f}s")
        self.running = False
        self.run_button.configure(state=NORMAL)

if __name__ == '__main__':
    root = Tk()
    app = RenamerApp(root)
    root.mainloop()
