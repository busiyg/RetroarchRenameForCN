# 中文Rom助手（RetroarchRenameForCN）
一个为中文Rom快速在Retroarch模拟器,萤火虫系统中获取游戏封面的方案

## 功能
- 把不规范中文名用模糊匹配翻译成官方英文名，在retroarch中就可以自动下载封面
- 修改retroarch的lpl列表配置文件，将英文的label修改为标准中文名
- 修改萤火虫（knull）系统的列表配置文件，将英文的label修改为标准中文名
- 可以把修改后标准的英文名再改回标准中文名
![Screenshot](Screenshot.png)

## 运行
- 视频教程：[B站](https://www.bilibili.com/video/BV1oXWxzLEGi)
- 可以直接从[Releases](https://github.com/busiyg/RetroarchRenameForCN/releases)中下载打包好的exe，也可以自己配置环境运行源码
```bash
pip install rapidfuzz pandas
python rom_rename_tool.py
```

## 平台
- FC,SFC,GB,GBC,GBA,NDS,3DS,New 3DS,Wii,Wii U,PS1,PSP,MD,DC

## 注意
- 不支持罗马数字，请改为阿拉伯数字，比如最终幻想IV,请改为最终幻想4

## 致谢
- https://github.com/yingw/rom-name-cn 中文游戏名称数据库
- https://github.com/libretro/RetroArch 著名的全能模拟器
- https://github.com/knulli-cfw/distribution 萤火虫开源系统

## 给自己打个广告
- https://store.steampowered.com/app/2827280/_/ 一人开发的大富翁类独立游戏，26年登录Steam，感兴趣的小伙伴加个愿望单吧~
