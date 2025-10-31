"""
ROM Renamer - 图形界面主程序
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
from tkinter.ttk import Combobox
from rapidfuzz import fuzz

from config import APP_TITLE, DEFAULT_THRESHOLD, PLATFORM_CONFIG
from core import CSVMapper, FileNameCleaner, SmartMatcher, generate_unique_filename, is_chinese_filename


class RenamerApp:
    def __init__(self, master):
        self.master = master
        master.title(APP_TITLE)
        master.resizable(False, False)
        
        self.folder_var = StringVar()
        self.lpl_var = StringVar()
        self.xml_var = StringVar()
        self.threshold_var = IntVar(value=DEFAULT_THRESHOLD)
        self.platform_var = StringVar()
        self.mapper = CSVMapper()
        self.running = False
        
        self._build_ui()
    
    def _build_ui(self):
        """构建UI"""
        # 文件夹选择
        Label(self.master, text="ROM 文件夹：").grid(row=0, column=0, sticky='w', padx=6, pady=6)
        Entry(self.master, textvariable=self.folder_var, width=45).grid(row=0, column=1, padx=6, pady=6, columnspan=2)
        Button(self.master, text="浏览", command=lambda: self._browse(self.folder_var, True)).grid(row=0, column=3, padx=6)
        
        # 平台选择
        Label(self.master, text="平台：").grid(row=0, column=4, sticky='w', padx=(20, 6), pady=6)
        self.platform_combo = Combobox(self.master, textvariable=self.platform_var, 
                                       values=sorted(PLATFORM_CONFIG.keys()), 
                                       state='readonly', width=18)
        self.platform_combo.grid(row=0, column=5, padx=6, pady=6)
        self.platform_combo.set('')  # 默认为空
        
        # LPL选择
        Label(self.master, text="LPL 播放列表：").grid(row=1, column=0, sticky='w', padx=6, pady=6)
        Entry(self.master, textvariable=self.lpl_var, width=45).grid(row=1, column=1, padx=6, pady=6, columnspan=2)
        Button(self.master, text="浏览", command=lambda: self._browse(self.lpl_var, False, "lpl")).grid(row=1, column=3, padx=6)
        Button(self.master, text="转换LPL", command=self._start_lpl, width=10).grid(row=1, column=4, columnspan=2, padx=(20, 6), pady=6)
        
        # XML选择（萤火虫）
        Label(self.master, text="萤火虫列表：").grid(row=2, column=0, sticky='w', padx=6, pady=6)
        Entry(self.master, textvariable=self.xml_var, width=45).grid(row=2, column=1, padx=6, pady=6, columnspan=2)
        Button(self.master, text="浏览", command=lambda: self._browse(self.xml_var, False, "xml")).grid(row=2, column=3, padx=6)
        Button(self.master, text="转换XML", command=self._start_xml, width=10).grid(row=2, column=4, columnspan=2, padx=(20, 6), pady=6)
        
        # 阈值设置
        Label(self.master, text="匹配阈值 (0-100)：").grid(row=3, column=0, sticky='w', padx=6, pady=6)
        Entry(self.master, textvariable=self.threshold_var, width=8).grid(row=3, column=1, sticky='w', padx=6, pady=6)
        Label(self.master, text="(自动清理文件名前缀并匹配CSV)", fg="gray").grid(row=3, column=1, columnspan=3, sticky='e', padx=6)
        
        # 操作按钮
        self.preview_btn = Button(self.master, text="预览重命名效果", command=self._start_preview, width=18)
        self.preview_btn.grid(row=4, column=0, padx=6, pady=6)
        self.run_btn = Button(self.master, text="执行重命名", command=self._start_rename, width=15)
        self.run_btn.grid(row=4, column=1, sticky='w', padx=6, pady=6)
        Button(self.master, text="清空日志", command=self._clear_log, width=10).grid(row=4, column=3, padx=6, pady=6)
        
        # 日志区域
        Label(self.master, text="日志/进度：").grid(row=5, column=0, sticky='nw', padx=6, pady=6)
        self.log = ScrolledText(self.master, width=100, height=20, state=DISABLED)
        self.log.grid(row=5, column=1, columnspan=5, padx=6, pady=6)
        
        # 作者信息
        author_frame = Frame(self.master)
        author_frame.grid(row=6, column=1, columnspan=5, pady=6)
        Label(author_frame, text="作者：").pack(side='left')
        Button(author_frame, text="奇个旦", fg="blue", cursor="hand2", relief="flat",
               command=lambda: webbrowser.open("https://space.bilibili.com/332938511")).pack(side='left')
    
    def _browse(self, var, is_folder, file_type=None):
        """浏览文件/文件夹"""
        if is_folder:
            result = filedialog.askdirectory()
        else:
            if file_type == "lpl":
                result = filedialog.askopenfilename(filetypes=[("LPL files", "*.lpl")])
            elif file_type == "xml":
                result = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
            else:
                result = filedialog.askopenfilename()
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
        self.preview_btn.configure(state=DISABLED)
        self.running = True
        self.mapper.cache.clear()
        threading.Thread(target=callback, args=(*args, threshold), daemon=True).start()
    
    def _start_preview(self):
        """启动预览模式"""
        folder = self.folder_var.get().strip()
        platform = self.platform_var.get().strip()
        
        if not folder or not os.path.isdir(folder):
            self._log("错误：请选择有效的ROM文件夹")
            return
        if not platform:
            self._log("错误：请选择平台类型")
            return
        
        self._validate_and_start(self._preview_roms, folder, platform)
    
    def _start_rename(self):
        """启动ROM重命名"""
        folder = self.folder_var.get().strip()
        platform = self.platform_var.get().strip()
        
        if not folder or not os.path.isdir(folder):
            self._log("错误：请选择有效的ROM文件夹")
            return
        if not platform:
            self._log("错误：请选择平台类型")
            return
        
        self._validate_and_start(self._rename_roms, folder, platform)
    
    def _start_lpl(self):
        """启动LPL转换"""
        lpl_path = self.lpl_var.get().strip()
        if not lpl_path or not os.path.exists(lpl_path):
            self._log("错误：请选择有效的LPL文件")
            return
        self._validate_and_start(self._convert_lpl, lpl_path)
    
    def _start_xml(self):
        """启动XML转换"""
        xml_path = self.xml_var.get().strip()
        if not xml_path or not os.path.exists(xml_path):
            self._log("错误：请选择有效的XML文件")
            return
        self._validate_and_start(self._convert_xml, xml_path)
    
    def _preview_roms(self, folder, platform, threshold):
        """预览重命名效果（不实际修改）"""
        from time import time
        start = time()
        self._log("=" * 70)
        self._log(f"【预览模式】开始分析ROM文件... [平台: {platform}]")
        
        # 获取平台支持的扩展名
        valid_extensions = self.mapper.get_platform_extensions(platform)
        if not valid_extensions:
            self._log(f"错误：未找到平台 {platform} 的配置")
            self._finish()
            return
        
        # 获取CSV路径
        csv_path = self.mapper.get_csv_path(platform)
        if not csv_path:
            self._log(f"错误：未找到平台 {platform} 的CSV文件")
            self._finish()
            return
        
        stats = {'total': 0, 'will_rename': 0, 'skipped': 0, 'english': 0, 'wrong_ext': 0, 'errors': 0}
        
        for filename in os.listdir(folder):
            if not os.path.isfile(os.path.join(folder, filename)):
                continue
            
            stats['total'] += 1
            name, ext = os.path.splitext(filename)
            
            try:
                # 检查扩展名是否匹配
                if ext.lower() not in valid_extensions:
                    stats['wrong_ext'] += 1
                    continue
                
                # 跳过英文文件
                if not is_chinese_filename(name):
                    stats['english'] += 1
                    continue
                
                # 匹配预览
                mapping = self.mapper.load_mapping(csv_path)
                cleaned = FileNameCleaner.clean(name)
                match, score = SmartMatcher.match(cleaned, mapping['cn_list'], threshold)
                
                if match and (eng := mapping['cn_to_eng'].get(match)):
                    new_name = generate_unique_filename(folder, eng + ext)
                    stats['will_rename'] += 1
                    self._log(f"✓ [预览] {filename}\n  → {new_name}\n  [分数: {score:.1f}]")
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
        self._log(f"将跳过: {stats['skipped']} | 错误扩展名: {stats['wrong_ext']} | 错误: {stats['errors']}")
        self._log(f"\n支持的扩展名: {', '.join(valid_extensions)}")
        self._log("\n⚡ 提示：如果预览效果满意，点击「执行重命名」按钮正式重命名文件")
        self._finish()
    
    def _rename_roms(self, folder, platform, threshold):
        """重命名ROM文件"""
        from time import time
        start = time()
        self._log("=" * 70)
        self._log(f"开始重命名ROM文件... [平台: {platform}]")
        
        # 获取平台支持的扩展名
        valid_extensions = self.mapper.get_platform_extensions(platform)
        if not valid_extensions:
            self._log(f"错误：未找到平台 {platform} 的配置")
            self._finish()
            return
        
        # 获取CSV路径
        csv_path = self.mapper.get_csv_path(platform)
        if not csv_path:
            self._log(f"错误：未找到平台 {platform} 的CSV文件")
            self._finish()
            return
        
        stats = {'total': 0, 'renamed': 0, 'skipped': 0, 'english': 0, 'wrong_ext': 0, 'errors': 0}
        
        for filename in os.listdir(folder):
            if not os.path.isfile(os.path.join(folder, filename)):
                continue
            
            stats['total'] += 1
            name, ext = os.path.splitext(filename)
            
            try:
                # 检查扩展名是否匹配
                if ext.lower() not in valid_extensions:
                    stats['wrong_ext'] += 1
                    continue
                
                # 跳过英文文件
                if not is_chinese_filename(name):
                    stats['english'] += 1
                    continue
                
                # 匹配并重命名
                mapping = self.mapper.load_mapping(csv_path)
                cleaned = FileNameCleaner.clean(name)
                match, score = SmartMatcher.match(cleaned, mapping['cn_list'], threshold)
                
                if match and (eng := mapping['cn_to_eng'].get(match)):
                    new_name = generate_unique_filename(folder, eng + ext)
                    os.rename(os.path.join(folder, filename), os.path.join(folder, new_name))
                    stats['renamed'] += 1
                    self._log(f"✓ {filename}\n  → {new_name}\n  [分数: {score:.1f}]")
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
        self._log(f"未匹配: {stats['skipped']} | 错误扩展名: {stats['wrong_ext']} | 错误: {stats['errors']}")
        self._log(f"\n支持的扩展名: {', '.join(valid_extensions)}")
        self._log(f"祝你玩的开心！🎮")
        
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
            
            stats = {'total': len(lpl['items']), 'converted': 0, 'skipped': 0, 'no_match': 0}
            platform_stats = {}
            
            # 尝试从所有平台配置中匹配
            for item in lpl['items']:
                label = item.get('label', '')
                ext = Path(item.get('path', '')).suffix.lower()
                
                # 查找支持该扩展名的平台
                matched_platform = None
                for platform_name, config in PLATFORM_CONFIG.items():
                    if ext in config['extensions']:
                        matched_platform = platform_name
                        break
                
                if not matched_platform:
                    stats['no_match'] += 1
                    self._log(f"⚠ 未找到支持扩展名 {ext} 的平台: {label}")
                    continue
                
                csv_path = self.mapper.get_csv_path(matched_platform)
                if not csv_path:
                    stats['no_match'] += 1
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
                    self._log(f"✓ {label}\n  → {best_match} [{matched_platform}, {best_score:.1f}]")
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
            self._log(f"总计: {stats['total']} | 已转换: {stats['converted']} | 保持: {stats['skipped']} | 无匹配: {stats['no_match']}")
            self._log(f"已保存到桌面: {clean_name}")
            
            if platform_stats:
                self._log("\n使用的CSV统计:")
                for csv, count in sorted(platform_stats.items()):
                    self._log(f"  • {csv}: {count} 个条目")
        
        except Exception as e:
            self._log(f"✗ 处理失败: {e}")
        
        self._finish()
    
    def _convert_xml(self, xml_path, threshold):
        """转换萤火虫XML播放列表"""
        from time import time
        from core import parse_xml_playlist, save_xml_playlist
        
        start = time()
        self._log("=" * 70)
        self._log(f"开始转换萤火虫XML: {Path(xml_path).name}")
        
        try:
            tree, root = parse_xml_playlist(xml_path)
            
            stats = {'total': 0, 'converted': 0, 'skipped': 0, 'no_match': 0}
            platform_stats = {}
            
            for game in root.findall('game'):
                name_elem = game.find('name')
                path_elem = game.find('path')
                
                if name_elem is None or path_elem is None:
                    continue
                
                stats['total'] += 1
                label = name_elem.text or ''
                ext = Path(path_elem.text or '').suffix.lower()
                
                # 查找支持该扩展名的平台
                matched_platform = None
                for platform_name, config in PLATFORM_CONFIG.items():
                    if ext in config['extensions']:
                        matched_platform = platform_name
                        break
                
                if not matched_platform:
                    stats['no_match'] += 1
                    self._log(f"⚠ 未找到支持扩展名 {ext} 的平台: {label}")
                    continue
                
                csv_path = self.mapper.get_csv_path(matched_platform)
                if not csv_path:
                    stats['no_match'] += 1
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
                    name_elem.text = best_match
                    stats['converted'] += 1
                    self._log(f"✓ {label}\n  → {best_match} [{matched_platform}, {best_score:.1f}]")
                else:
                    stats['skipped'] += 1
                    self._log(f"⊙ 保持: {label} ({best_score:.1f})")
            
            # 保存到桌面
            desktop = Path.home() / "Desktop"
            save_path = desktop / save_name
            
            save_xml_playlist(tree, save_path)
            
            self._log("=" * 70)
            self._log(f"完成! 耗时: {time()-start:.1f}s")
            self._log(f"总计: {stats['total']} | 已转换: {stats['converted']} | 保持: {stats['skipped']} | 无匹配: {stats['no_match']}")
            self._log(f"已保存到桌面: {save_name}")
            
            if platform_stats:
                self._log("\n使用的CSV统计:")
                for csv, count in sorted(platform_stats.items()):
                    self._log(f"  • {csv}: {count} 个条目")
        
        except Exception as e:
            self._log(f"✗ 处理失败: {e}")
        
        self._finish()
    
    def _finish(self):
        """完成任务"""
        self.running = False
        self.run_btn.configure(state=NORMAL)
        self.preview_btn.configure(state=NORMAL)


if __name__ == '__main__':
    root = Tk()
    app = RenamerApp(root)
    root.mainloop()