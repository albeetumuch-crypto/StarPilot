#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量重新命名照片檔案
將照片改名為：旅遊_001.jpg、旅遊_002.jpg 等格式
"""

import os
import sys
from pathlib import Path


def preview_rename(source_dir, prefix="旅遊", dry_run=True):
    """預覽重新命名結果"""
    source_path = Path(source_dir)

    if not source_path.exists():
        print(f"❌ 錯誤：資料夾不存在：{source_dir}")
        return False

    # 取得所有圖片檔案
    image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'}
    files = [f for f in source_path.iterdir()
             if f.is_file() and f.suffix.lower() in image_extensions]

    if not files:
        print(f"❌ 錯誤：找不到圖片檔案在 {source_dir}")
        return False

    # 按檔名排序
    files = sorted(files)

    print(f"\n📋 預覽重新命名結果 ({len(files)} 個檔案)")
    print("=" * 70)

    changes = []
    for idx, old_file in enumerate(files, 1):
        new_filename = f"{prefix}_{idx:03d}{old_file.suffix}"
        new_path = old_file.parent / new_filename

        changes.append({
            'old': old_file.name,
            'new': new_filename,
            'old_path': old_file,
            'new_path': new_path
        })

        print(f"{idx:2d}. {old_file.name:30s} → {new_filename}")

    print("=" * 70)

    if not dry_run:
        print("\n⏳ 正在執行重新命名...")
        for change in changes:
            try:
                change['old_path'].rename(change['new_path'])
                print(f"✅ {change['old']} → {change['new']}")
            except Exception as e:
                print(f"❌ 失敗：{change['old']} ({str(e)})")
        print(f"\n✨ 完成！已重新命名 {len(files)} 個檔案")

    return True


if __name__ == "__main__":
    source_dir = "/workspaces/StarPilot/examples/03_批次處理/測試資料/待重新命名"

    # 首先預覽
    print("\n🔍 第一步：預覽重新命名結果")
    preview_rename(source_dir, prefix="旅遊", dry_run=True)

    # 確認執行
    print("\n" + "=" * 70)
    response = input("確認執行重新命名？(輸入 'yes' 確認): ").strip().lower()

    if response == 'yes':
        preview_rename(source_dir, prefix="旅遊", dry_run=False)
    else:
        print("❌ 已取消操作")
