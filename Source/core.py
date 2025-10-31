"""
ROM Renamer - 核心功能模块
包含CSV映射、文件名清理、智能匹配等核心功能
"""
import os
import re
from pathlib import Path
import pandas as pd
from rapidfuzz import process, fuzz
from config import PLATFORM_CONFIG, CSV_ROOT_DIR, MATCH_WEIGHTS, LENGTH_RATIO_POWER, SUBSTRING_PENALTY


class CSVMapper:
    """CSV映射和缓存管理"""
    def __init__(self, csv_root=CSV_ROOT_DIR):
        self.csv_dir = Path(__file__).parent / csv_root
        self.cache = {}
    
    def get_csv_path(self, platform_name):
        """根据平台名获取CSV路径"""
        if platform_name not in PLATFORM_CONFIG:
            return None
        csv_name = PLATFORM_CONFIG[platform_name]['csv']
        csv_path = self.csv_dir / csv_name
        return csv_path if csv_path.exists() else None
    
    def get_platform_extensions(self, platform_name):
        """获取平台支持的扩展名列表"""
        if platform_name not in PLATFORM_CONFIG:
            return []
        return PLATFORM_CONFIG[platform_name]['extensions']
    
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
            
            # 加权计算
            weights = [
                MATCH_WEIGHTS['token_set_ratio'],
                MATCH_WEIGHTS['ratio'],
                MATCH_WEIGHTS['partial_ratio'],
                MATCH_WEIGHTS['token_sort_ratio']
            ]
            composite = sum(s * w for s, w in zip(scores, weights))
            
            # 长度惩罚
            len_ratio = min(len(query), len(candidate)) / max(len(query), len(candidate))
            composite *= (len_ratio ** LENGTH_RATIO_POWER)
            
            # 子串惩罚
            if len(candidate) < len(query) and candidate in query:
                composite *= SUBSTRING_PENALTY
            
            if composite > best_score:
                best_score, best_match = composite, candidate
        
        return (best_match, best_score) if best_score >= threshold else (None, best_score)


def generate_unique_filename(folder, filename):
    """生成唯一文件名（避免重复）"""
    base, ext = os.path.splitext(filename)
    candidate, i = filename, 1
    while os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{base} ({i}){ext}"
        i += 1
    return candidate


def is_chinese_filename(name):
    """检查文件名是否包含中文"""
    return bool(re.search(r'[\u4e00-\u9fff]', name))


def parse_xml_playlist(xml_path):
    """解析XML播放列表（萤火虫格式）"""
    import xml.etree.ElementTree as ET
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        if root.tag != 'gameList':
            raise ValueError("不是有效的gamelist.xml格式")
        return tree, root
    except Exception as e:
        raise ValueError(f"解析XML失败: {e}")


def save_xml_playlist(tree, save_path):
    """保存XML播放列表"""
    import xml.etree.ElementTree as ET
    # 保持原有格式
    ET.indent(tree, space="\t")
    tree.write(save_path, encoding='utf-8', xml_declaration=True)