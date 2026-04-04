#!/usr/bin/env python3
"""
core_engine.py - 图片转CDN链接生成器，带Git自动化功能
扫描images/文件夹，生成jsDelivr CDN链接，导出到Excel，并运行git命令。
"""

import os
import sys
import subprocess
from pathlib import Path
from typing import List, Tuple

# 设置标准输出编码为UTF-8，避免Windows控制台乱码
import io
if sys.stdout.encoding is None or sys.stdout.encoding.upper() not in ['UTF-8', 'UTF8']:
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    except Exception:
        pass  # 如果失败，继续使用默认编码

# 支持的图片文件扩展名
支持的图片扩展名 = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp', '.svg')

def 获取github信息() -> Tuple[str, str]:
    """
    从git远程origin提取GitHub用户名和仓库名。
    返回 (用户名, 仓库名)
    """
    try:
        # 获取远程origin URL
        结果 = subprocess.run(['git', 'remote', 'get-url', 'origin'],
                              capture_output=True, text=True, check=True)
        远程url = 结果.stdout.strip()

        # 解析URL
        if 远程url.startswith('https://github.com/'):
            # HTTPS格式: https://github.com/用户名/仓库名.git
            路径部分 = 远程url.rstrip('.git').split('/')
            用户名 = 路径部分[-2]
            仓库名 = 路径部分[-1]
        elif 远程url.startswith('git@github.com:'):
            # SSH格式: git@github.com:用户名/仓库名.git
            路径部分 = 远程url.split(':')[1].rstrip('.git').split('/')
            用户名 = 路径部分[0]
            仓库名 = 路径部分[1]
        else:
            print(f"未知的git远程格式: {远程url}")
            print("请设置git远程: git remote add origin <你的仓库URL>")
            sys.exit(1)

        return 用户名, 仓库名
    except subprocess.CalledProcessError:
        print("错误: 未找到git远程 'origin'。")
        print("请设置git远程: git remote add origin <你的仓库URL>")
        sys.exit(1)
    except Exception as 错误:
        print(f"解析git远程时出错: {错误}")
        sys.exit(1)

def 扫描图片(图片目录: str) -> List[Tuple[str, str]]:
    """
    扫描图片目录中的图片文件。
    返回列表，每个元素是(文件名, 相对路径)
    """
    图片文件列表 = []

    # 确保图片目录存在
    if not os.path.exists(图片目录):
        print(f"警告: 图片目录 '{图片目录}' 不存在。")
        return 图片文件列表

    # 遍历目录
    for 根目录, 子目录列表, 文件列表 in os.walk(图片目录):
        for 文件名 in 文件列表:
            if 文件名.lower().endswith(支持的图片扩展名):
                完整路径 = os.path.join(根目录, 文件名)
                # 获取相对于仓库根目录的相对路径
                相对路径 = os.path.relpath(完整路径, start=os.getcwd())
                图片文件列表.append((文件名, 相对路径))

    return 图片文件列表

def 生成cdn链接(图片文件列表: List[Tuple[str, str]], 用户名: str, 仓库名: str) -> List[Tuple[str, str, str]]:
    """
    为图片文件生成jsDelivr CDN链接。
    返回列表，每个元素是(文件名, 相对路径, CDN链接)
    """
    cdn链接列表 = []

    for 文件名, 相对路径 in 图片文件列表:
        # 将Windows反斜杠转换为URL使用的正斜杠
        url路径 = 相对路径.replace('\\', '/')
        cdn链接 = f"https://cdn.jsdelivr.net/gh/{用户名}/{仓库名}/{url路径}"
        cdn链接列表.append((文件名, 相对路径, cdn链接))

    return cdn链接列表

def 导出到excel(数据: List[Tuple[str, str, str]], 输出目录: str):
    """
    将数据导出到输出目录中的Excel文件。
    """
    try:
        import pandas as pd
    except ImportError:
        print("错误: 需要pandas库来导出Excel。")
        print("请安装: pip install pandas openpyxl")
        sys.exit(1)

    # 确保输出目录存在
    os.makedirs(输出目录, exist_ok=True)

    # 创建DataFrame
    数据框 = pd.DataFrame(数据, columns=['文件名', '相对路径', 'CDN 链接'])

    # 生成带时间戳的输出文件名
    from datetime import datetime
    时间戳 = datetime.now().strftime('%Y%m%d_%H%M%S')
    输出文件 = os.path.join(输出目录, f'cdn链接_{时间戳}.xlsx')

    # 导出到Excel
    try:
        数据框.to_excel(输出文件, index=False, engine='openpyxl')
        print(f"[OK] Excel文件已导出: {输出文件}")
        return 输出文件
    except Exception as 错误:
        print(f"导出到Excel时出错: {错误}")
        sys.exit(1)

def 运行git命令():
    """
    执行git add、commit和push命令。
    """
    命令列表 = [
        ['git', 'add', '.'],
        ['git', 'commit', '-m', 'Update'],
        ['git', 'push', 'origin', 'main']
    ]

    for 命令 in 命令列表:
        try:
            print(f"正在执行: {' '.join(命令)}")
            结果 = subprocess.run(命令, capture_output=True, text=True, check=True)
            print(结果.stdout)
            if 结果.stderr:
                print(f"标准错误: {结果.stderr}")
        except subprocess.CalledProcessError as 错误:
            print(f"执行命令时出错: {' '.join(命令)}")
            print(f"标准错误: {错误.stderr}")
            # 即使一个命令失败，也继续执行下一个命令
            continue
        except Exception as 错误:
            print(f"意外错误: {错误}")

def 主函数():
    print("=" * 60)
    print("图片转CDN链接生成器")
    print("=" * 60)

    # 获取当前工作目录（仓库根目录）
    仓库根目录 = os.getcwd()
    print(f"仓库根目录: {仓库根目录}")

    # 定义目录路径
    图片目录 = os.path.join(仓库根目录, 'images')
    输出目录 = os.path.join(仓库根目录, 'output')

    # 步骤1: 获取GitHub用户名和仓库名
    print("\n[1] 正在获取GitHub仓库信息...")
    用户名, 仓库名 = 获取github信息()
    print(f"   用户名: {用户名}")
    print(f"   仓库名: {仓库名}")

    # 步骤2: 扫描图片
    print("\n[2] 正在扫描图片目录...")
    图片文件列表 = 扫描图片(图片目录)

    if not 图片文件列表:
        print("   在images/目录中未找到图片文件。")
        print("   请将图片添加到images/文件夹中，然后重试。")
        return

    print(f"   找到 {len(图片文件列表)} 个图片文件。")

    # 步骤3: 生成CDN链接
    print("\n[3] 正在生成CDN链接...")
    cdn数据 = 生成cdn链接(图片文件列表, 用户名, 仓库名)

    for 文件名, _, cdn链接 in cdn数据[:3]:  # 显示前3个作为预览
        print(f"   {文件名} → {cdn链接}")
    if len(cdn数据) > 3:
        print(f"   ... 还有 {len(cdn数据) - 3} 个")

    # 步骤4: 导出到Excel
    print("\n[4] 正在导出到Excel...")
    excel文件 = 导出到excel(cdn数据, 输出目录)

    # 步骤5: 运行git命令
    print("\n[5] 正在运行git命令...")
    运行git命令()

    print("\n" + "=" * 60)
    print("处理完成！")
    print(f"Excel文件: {excel文件}")
    print("=" * 60)

if __name__ == "__main__":
    主函数()