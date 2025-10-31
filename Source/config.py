"""
ROM Renamer - 配置文件
管理CSV映射和全局配置
"""

# 应用配置
APP_TITLE = "中文ROM助手 1.3"
CSV_ROOT_DIR = 'rom-name-cn-master'  # CSV文件根目录
DEFAULT_THRESHOLD = 40  # 默认匹配阈值

# 平台配置：平台名 -> (CSV文件名, 支持的扩展名列表)
PLATFORM_CONFIG = {
    'Game Boy Advance': {
        'csv': 'Nintendo - Game Boy Advance.csv',
        'extensions': ['.gba', '.zip']
    },
    'Game Boy Color': {
        'csv': 'Nintendo - Game Boy Color.csv',
        'extensions': ['.gbc', '.zip']
    },
    'Game Boy': {
        'csv': 'Nintendo - Game Boy.csv',
        'extensions': ['.gb', '.zip']
    },
    'Nintendo 3DS': {
        'csv': 'Nintendo - New Nintendo 3DS.csv',
        'extensions': ['.3ds', '.cia', '.zip']
    },
    'Nintendo DS': {
        'csv': 'Nintendo - Nintendo DS.csv',
        'extensions': ['.nds', '.zip']
    },
    'Nintendo 64': {
        'csv': 'Nintendo - Nintendo 64.csv',
        'extensions': ['.n64', '.z64', '.v64', '.zip']
    },
    'NES': {
        'csv': 'Nintendo - Nintendo Entertainment System.csv',
        'extensions': ['.nes', '.zip']
    },
    'Super Nintendo': {
        'csv': 'Nintendo - Super Nintendo Entertainment System.csv',
        'extensions': ['.sfc', '.smc', '.zip']
    },
    'Wii U': {
        'csv': 'Nintendo - Wii U.csv',
        'extensions': ['.wud', '.wux']
    },
    'Wii': {
        'csv': 'Nintendo - Wii.csv',
        'extensions': ['.wbfs', '.iso', '.wbf', '.rvz']
    },
    'PlayStation Portable': {
        'csv': 'Sony - PlayStation Portable.csv',
        'extensions': ['.cso', '.iso', '.zip']
    },
    'PlayStation': {
        'csv': 'Sony - PlayStation.csv',
        'extensions': ['.bin', '.cue', '.img', '.mdf', '.pbp', '.toc', '.cbn', '.m3u']
    },
    'Dreamcast': {
        'csv': 'Sega - Dreamcast.csv',
        'extensions': ['.cdi', '.gdi', '.chd']
    },
    'Mega Drive': {
        'csv': 'Sega - Mega Drive - Genesis.csv',
        'extensions': ['.md', '.gen', '.smd', '.bin', '.zip']
    },
}

# 匹配算法权重配置
MATCH_WEIGHTS = {
    'token_set_ratio': 0.35,
    'ratio': 0.25,
    'partial_ratio': 0.25,
    'token_sort_ratio': 0.15
}

# 长度比例权重
LENGTH_RATIO_POWER = 0.5

# 子串惩罚系数
SUBSTRING_PENALTY = 0.7