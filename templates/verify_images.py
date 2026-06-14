"""
中期考核报告图片验证脚本 v1.0

检查图片是否满足报告要求：
- 图片数量 >= 15
- 图片格式有效（PNG/JPG/JPEG/TIFF/BMP）
- 图片尺寸合理（宽度 >= 400px）
- 图片编号连续无跳号

使用方法：
    E:/Anaconda/python.exe verify_images.py [工作目录]

环境要求：
    Python >= 3.7
    依赖：Pillow
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import NamedTuple

try:
    from PIL import Image
except ImportError:
    print('错误：需要安装 Pillow 库')
    print('运行: pip install Pillow')
    sys.exit(1)


class ImageInfo(NamedTuple):
    """图片信息"""
    path: str
    name: str
    width: int
    height: int
    format: str
    size_kb: float


VALID_FORMATS: set[str] = {'PNG', 'JPG', 'JPEG', 'TIFF', 'BMP'}
MIN_IMAGE_COUNT: int = 15
MIN_WIDTH: int = 400
MIN_SIZE_KB: float = 5.0


def scan_images(directory: str) -> list[ImageInfo]:
    """扫描目录中的所有图片文件"""
    images: list[ImageInfo] = []
    if not os.path.isdir(directory):
        return images

    for filename in sorted(os.listdir(directory)):
        filepath = os.path.join(directory, filename)
        if not os.path.isfile(filepath):
            continue

        ext = filename.rsplit('.', 1)[-1].upper() if '.' in filename else ''
        if ext not in VALID_FORMATS:
            continue

        try:
            with Image.open(filepath) as img:
                width, height = img.size
                format_str = img.format or ext
                size_kb = os.path.getsize(filepath) / 1024
                images.append(ImageInfo(
                    path=filepath,
                    name=filename,
                    width=width,
                    height=height,
                    format=format_str,
                    size_kb=size_kb,
                ))
        except Exception as e:
            print(f'  警告：无法读取图片 {filename}: {e}')

    return images


def check_numbering(images: list[ImageInfo]) -> list[str]:
    """检查图片编号是否连续（图1, 图2, ...）"""
    issues: list[str] = []
    numbered: list[int] = []

    for img in images:
        name = img.name
        # 尝试从文件名提取编号：图1.png, fig1.png, 1.png 等
        for prefix in ['图', 'fig', 'Fig', 'FIG', '']:
            if name.startswith(prefix) or (prefix == '' and name[0].isdigit()):
                num_str = name[len(prefix):].split('.')[0]
                try:
                    numbered.append(int(num_str))
                except ValueError:
                    pass
                break

    if not numbered:
        return issues

    numbered.sort()
    for i in range(len(numbered) - 1):
        if numbered[i + 1] - numbered[i] > 1:
            gap = list(range(numbered[i] + 1, numbered[i + 1]))
            issues.append(f'图片编号跳号：缺少 图{gap[0]} 到 图{gap[-1]}')

    return issues


def verify_images(base_dir: str) -> bool:
    """验证图片是否满足报告要求

    Returns:
        True if all checks pass, False otherwise
    """
    print(f'=== 中期考核报告图片验证 ===')
    print(f'工作目录: {base_dir}\n')

    all_passed = True

    # 1. 扫描各目录
    dirs_to_check = [
        ('图片', '用户提供的图片'),
        ('ppt_images', 'PPT提取的图片'),
        ('paper_images', '论文提取的图片'),
    ]

    all_images: list[ImageInfo] = []
    for dirname, desc in dirs_to_check:
        dirpath = os.path.join(base_dir, dirname)
        images = scan_images(dirpath)
        if images:
            print(f'[{desc}] {dirpath}')
            print(f'  找到 {len(images)} 张图片')
            all_images.extend(images)
        else:
            print(f'[{desc}] {dirpath} — 未找到图片')

    print(f'\n总计: {len(all_images)} 张图片\n')

    # 2. 检查图片数量
    print('--- 检查图片数量 ---')
    if len(all_images) >= MIN_IMAGE_COUNT:
        print(f'  ✓ 图片数量 {len(all_images)} >= {MIN_IMAGE_COUNT}')
    else:
        print(f'  ✗ 图片数量 {len(all_images)} < {MIN_IMAGE_COUNT}（不足）')
        all_passed = False

    # 3. 检查图片格式
    print('\n--- 检查图片格式 ---')
    format_ok = True
    for img in all_images:
        if img.format not in VALID_FORMATS:
            print(f'  ✗ {img.name}: 格式 {img.format} 不支持')
            format_ok = False
            all_passed = False
    if format_ok:
        print(f'  ✓ 所有图片格式有效')

    # 4. 检查图片尺寸
    print('\n--- 检查图片尺寸 ---')
    size_ok = True
    for img in all_images:
        if img.width < MIN_WIDTH:
            print(f'  ✗ {img.name}: 宽度 {img.width}px < {MIN_WIDTH}px（过小）')
            size_ok = False
            all_passed = False
    if size_ok:
        print(f'  ✓ 所有图片宽度 >= {MIN_WIDTH}px')

    # 5. 检查图片大小
    print('\n--- 检查图片文件大小 ---')
    size_kb_ok = True
    for img in all_images:
        if img.size_kb < MIN_SIZE_KB:
            print(f'  ✗ {img.name}: 文件大小 {img.size_kb:.1f}KB < {MIN_SIZE_KB}KB（可能损坏）')
            size_kb_ok = False
            all_passed = False
    if size_kb_ok:
        print(f'  ✓ 所有图片文件大小 >= {MIN_SIZE_KB}KB')

    # 6. 检查编号连续性
    print('\n--- 检查图片编号连续性 ---')
    numbering_issues = check_numbering(all_images)
    if numbering_issues:
        for issue in numbering_issues:
            print(f'  ✗ {issue}')
        all_passed = False
    else:
        print(f'  ✓ 图片编号连续（或无编号命名）')

    # 7. 图片详情
    print('\n--- 图片详情 ---')
    print(f'{"文件名":<30} {"尺寸":<15} {"格式":<8} {"大小":<10}')
    print('-' * 65)
    for img in all_images:
        print(f'{img.name:<30} {img.width}x{img.height:<10} {img.format:<8} {img.size_kb:.1f}KB')

    # 总结
    print('\n=== 验证结果 ===')
    if all_passed:
        print('✓ 所有检查通过！图片满足报告要求。')
    else:
        print('✗ 部分检查未通过，请根据上述提示修复。')

    return all_passed


def main() -> None:
    if len(sys.argv) > 1:
        base_dir = sys.argv[1]
    else:
        base_dir = os.getcwd()

    if not os.path.isdir(base_dir):
        print(f'错误：目录不存在: {base_dir}')
        sys.exit(1)

    success = verify_images(base_dir)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
