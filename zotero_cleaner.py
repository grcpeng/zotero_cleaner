# -*- coding: utf-8 -*-
"""
Zotero 完整清理工具
1. 清理重复的 PDF 文件
2. 清理孤立的 PDF 文件（不在数据库中）
3. 删除空文件夹
4. 删除不含 PDF 的文件夹
"""
import os
import sys
import re
import shutil
import sqlite3
import stat
import configparser
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from collections import defaultdict
import pandas as pd


def get_zotero_dirs():
    """获取 Zotero 数据目录"""
    profile_dirs = {
        'darwin': Path.home() / 'Library/Application Support/Zotero',
        'linux': Path.home() / '.zotero/zotero',
        'linux2': Path.home() / '.zotero/zotero',
        'win32': Path.home() / 'AppData/Roaming/Zotero/Zotero'
    }
    
    if sys.platform not in profile_dirs:
        print(f"错误：不支持的操作系统 {sys.platform}")
        sys.exit(1)
    
    profile_dir = profile_dirs[sys.platform]
    
    if not profile_dir.exists():
        print(f"错误：Zotero 配置目录不存在: {profile_dir}")
        sys.exit(1)
    
    config = configparser.ConfigParser()
    profiles_ini = profile_dir / 'profiles.ini'
    
    if not profiles_ini.exists():
        print(f"错误：找不到 profiles.ini: {profiles_ini}")
        sys.exit(1)
    
    config.read(str(profiles_ini), encoding='utf-8')
    prefs_js = profile_dir / config['Profile0']['Path'] / 'prefs.js'
    
    if not prefs_js.exists():
        print(f"错误：找不到 prefs.js: {prefs_js}")
        sys.exit(1)
    
    configs = prefs_js.read_text(encoding='utf-8')
    
    zotero_data_pat = re.compile(
        r'user_pref\("extensions\.zotero\.dataDir",\s*"(?P<zotero_data>[^"]+)"\);')
    zotero_match = zotero_data_pat.search(configs)
    
    if zotero_match:
        data_dir_str = zotero_match.group('zotero_data').replace('\\\\', '\\')
        zotero_data_dir = Path(data_dir_str)
    else:
        zotero_data_dir = profile_dir / 'Zotero'
        print(f"警告：未找到自定义数据目录，使用默认位置")
    
    storage_dir = zotero_data_dir / 'storage'
    
    if not zotero_data_dir.exists():
        print(f"错误：Zotero 数据目录不存在: {zotero_data_dir}")
        sys.exit(1)
    
    if not storage_dir.exists():
        print(f"错误：Storage 目录不存在: {storage_dir}")
        sys.exit(1)
    
    print(f"Zotero 数据目录: {zotero_data_dir}")
    print(f"Storage 目录: {storage_dir}")
    
    return zotero_data_dir, storage_dir


def collect_pdf_files(storage_dir):
    """收集 storage 目录中的所有 PDF 文件"""
    pdf_files = []
    
    print("\n正在扫描 PDF 文件...")
    for root, dirs, files in os.walk(storage_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                full_path = os.path.join(root, file)
                pdf_files.append((full_path, file))
    
    print(f"找到 {len(pdf_files)} 个 PDF 文件")
    return pdf_files


def get_database_pdfs(zotero_data_dir):
    """从 Zotero 数据库中获取所有 PDF 文件名和对应的文件夹"""
    db_path = zotero_data_dir / 'zotero.sqlite'
    
    if not db_path.exists():
        print(f"错误：数据库文件不存在: {db_path}")
        sys.exit(1)
    
    print("\n正在读取数据库...")
    try:
        with sqlite3.connect(str(db_path)) as con:
            query = """
            SELECT 
                ia.itemID,
                ia.parentItemID,
                ia.path,
                i.key as itemKey
            FROM itemAttachments ia
            LEFT JOIN items i ON ia.itemID = i.itemID
            WHERE ia.path IS NOT NULL
            """
            item_att = pd.read_sql_query(query, con=con)
            
            print(f"数据库中有 {len(item_att)} 条附件记录")
            
            db_files = {}
            db_folders = set()  # 数据库中所有有效的文件夹
            
            for _, row in item_att.iterrows():
                path = row['path']
                if isinstance(path, str) and ':' in path:
                    parts = path.split(':', 1)[1]
                    
                    if '/' in parts or '\\' in parts:
                        parts = parts.replace('\\', '/')
                        folder = parts.split('/')[0]
                        filename = parts.split('/')[-1]
                    else:
                        folder = row['itemKey']
                        filename = parts
                    
                    db_folders.add(folder)
                    
                    if filename not in db_files:
                        db_files[filename] = []
                    
                    db_files[filename].append({
                        'folder': folder,
                        'itemKey': row['itemKey'],
                        'db_path': path,
                        'itemID': row['itemID']
                    })
            
            print(f"数据库中有 {len(db_folders)} 个有效文件夹")
            return db_files, db_folders
    except Exception as e:
        print(f"错误：读取数据库失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


def clean_duplicate_pdfs(pdf_files, db_files, back_dir):
    """清理重复的 PDF 文件"""
    pdf_by_name = defaultdict(list)
    
    for full_path, filename in pdf_files:
        folder = os.path.basename(os.path.dirname(full_path))
        pdf_by_name[filename].append({
            'full_path': full_path,
            'folder': folder
        })
    
    duplicates = {name: files for name, files in pdf_by_name.items() if len(files) > 1}
    
    if not duplicates:
        print("\n没有发现重复的 PDF 文件")
        return 0
    
    print(f"\n{'='*60}")
    print(f"发现 {len(duplicates)} 个重复的文件名:")
    print(f"{'='*60}")
    
    files_to_delete = []
    
    for i, (filename, file_list) in enumerate(duplicates.items(), 1):
        print(f"\n{i}. {filename} (共 {len(file_list)} 份)")
        
        if filename in db_files:
            db_folders = [rec['folder'] for rec in db_files[filename]]
            print(f"   数据库中的文件夹: {db_folders}")
            
            for file_info in file_list:
                folder = file_info['folder']
                full_path = file_info['full_path']
                
                if folder in db_folders:
                    print(f"   ✓ 保留: {folder}/")
                else:
                    print(f"   ✗ 删除: {folder}/")
                    files_to_delete.append(full_path)
        else:
            print(f"   ⚠ 文件不在数据库中，全部标记为删除")
            for file_info in file_list:
                print(f"   ✗ 删除: {file_info['folder']}/")
                files_to_delete.append(file_info['full_path'])
    
    print(f"\n{'='*60}")
    print(f"总计需要删除 {len(files_to_delete)} 个重复文件")
    print(f"{'='*60}")
    
    if not files_to_delete:
        return 0
    
    choice = input("\n是否将这些重复文件移动到备份目录? (y/n): ").strip().lower()
    
    if choice != 'y':
        print("已取消操作")
        return 0
    
    print("\n开始移动重复文件...")
    success_count = 0
    
    for full_path in files_to_delete:
        try:
            filename = os.path.basename(full_path)
            folder = os.path.basename(os.path.dirname(full_path))
            
            dest_filename = f"dup_{folder}_{filename}"
            dest_path = os.path.join(back_dir, dest_filename)
            
            counter = 1
            while os.path.exists(dest_path):
                base, ext = os.path.splitext(dest_filename)
                dest_path = os.path.join(back_dir, f"{base}_{counter}{ext}")
                counter += 1
            
            shutil.move(full_path, dest_path)
            print(f"✓ 已移动: {folder}/{filename}")
            success_count += 1
        except Exception as e:
            print(f"✗ 移动失败 {filename}: {e}")
    
    print(f"\n成功移动 {success_count} 个重复文件")
    return success_count


def clean_orphaned_pdfs(pdf_files, db_files, back_dir):
    """清理孤立的 PDF 文件"""
    pdf_by_name = defaultdict(list)
    
    for full_path, filename in pdf_files:
        folder = os.path.basename(os.path.dirname(full_path))
        pdf_by_name[filename].append({
            'full_path': full_path,
            'folder': folder
        })
    
    orphaned = {name: files for name, files in pdf_by_name.items() if name not in db_files}
    
    if not orphaned:
        print("\n没有发现孤立的 PDF 文件")
        return 0
    
    print(f"\n{'='*60}")
    print(f"发现 {len(orphaned)} 个孤立文件（不在数据库中）:")
    print(f"{'='*60}")
    
    orphaned_files = []
    for filename, file_list in orphaned.items():
        for file_info in file_list:
            orphaned_files.append((file_info['full_path'], filename, file_info['folder']))
            print(f"  - {file_info['folder']}/{filename}")
    
    choice = input(f"\n是否将这 {len(orphaned_files)} 个孤立文件移动到备份目录? (y/n): ").strip().lower()
    
    if choice != 'y':
        print("已取消操作")
        return 0
    
    print("\n开始移动孤立文件...")
    success_count = 0
    
    for full_path, filename, folder in orphaned_files:
        try:
            dest_filename = f"orphan_{folder}_{filename}"
            dest_path = os.path.join(back_dir, dest_filename)
            
            counter = 1
            while os.path.exists(dest_path):
                base, ext = os.path.splitext(dest_filename)
                dest_path = os.path.join(back_dir, f"{base}_{counter}{ext}")
                counter += 1
            
            shutil.move(full_path, dest_path)
            print(f"✓ 已移动: {folder}/{filename}")
            success_count += 1
        except Exception as e:
            print(f"✗ 移动失败 {filename}: {e}")
    
    print(f"\n成功移动 {success_count} 个孤立文件")
    return success_count


def is_folder_empty(path):
    """检查文件夹是否为空"""
    try:
        items = os.listdir(path)
        items = [item for item in items if item not in ['.DS_Store', 'Thumbs.db', 'desktop.ini', '.zotero-ft-cache']]
        return len(items) == 0
    except (OSError, PermissionError):
        return False


def has_pdf_files(path):
    """检查文件夹是否包含 PDF 文件"""
    try:
        for root, dirs, files in os.walk(path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    return True
        return False
    except (OSError, PermissionError):
        return True


def remove_readonly(path):
    """移除只读属性"""
    try:
        os.chmod(path, stat.S_IWRITE)
    except:
        pass


def delete_folder_safe(dir_path):
    """安全删除文件夹"""
    try:
        remove_readonly(dir_path)
        
        for root, dirs, files in os.walk(dir_path):
            for d in dirs:
                remove_readonly(os.path.join(root, d))
            for f in files:
                remove_readonly(os.path.join(root, f))
        
        shutil.rmtree(dir_path)
        return True
    except Exception as e:
        print(f"⚠ 无法删除: {dir_path} - {e}")
        return False


def clean_empty_folders(storage_dir, db_folders):
    """清理空文件夹和不在数据库中的文件夹"""
    print(f"\n{'='*60}")
    print("清理空文件夹和无效文件夹")
    print(f"{'='*60}")
    
    empty_count = 0
    invalid_count = 0
    deleted = True
    
    while deleted:
        deleted = False
        folders_to_check = []
        
        # 收集所有子文件夹
        for root, dirs, files in os.walk(storage_dir, topdown=False):
            for dir_name in dirs:
                dir_path = os.path.join(root, dir_name)
                if os.path.abspath(dir_path) != os.path.abspath(storage_dir):
                    folders_to_check.append(dir_path)
        
        for dir_path in folders_to_check:
            try:
                folder_name = os.path.basename(dir_path)
                
                # 检查是否为空
                if is_folder_empty(dir_path):
                    os.rmdir(dir_path)
                    print(f"✓ 已删除空文件夹: {folder_name}")
                    empty_count += 1
                    deleted = True
                # 检查是否不含 PDF 且不在数据库中
                elif not has_pdf_files(dir_path) and folder_name not in db_folders:
                    if delete_folder_safe(dir_path):
                        print(f"✓ 已删除无效文件夹: {folder_name} (无PDF且不在数据库中)")
                        invalid_count += 1
                        deleted = True
            except PermissionError:
                print(f"⚠ 权限不足，跳过: {folder_name}")
            except Exception as e:
                print(f"⚠ 处理失败: {folder_name} - {e}")
    
    print(f"\n清理统计:")
    print(f"  - 删除空文件夹: {empty_count} 个")
    print(f"  - 删除无效文件夹: {invalid_count} 个")
    print(f"  - 总计: {empty_count + invalid_count} 个")
    
    return empty_count + invalid_count


def main():
    print("=" * 60)
    print("Zotero 完整清理工具")
    print("=" * 60)
    
    # 选择备份目录
    root = tk.Tk()
    root.withdraw()
    back_dir = filedialog.askdirectory(title='请选择备份目录（用于存放删除的文件）')
    
    if not back_dir:
        print("未选择备份目录，程序退出")
        sys.exit(0)
    
    print(f"\n备份目录: {back_dir}")
    
    if not os.path.exists(back_dir):
        os.makedirs(back_dir)
    
    # 获取 Zotero 目录
    zotero_data_dir, storage_dir = get_zotero_dirs()
    
    # 收集 PDF 文件
    pdf_files = collect_pdf_files(storage_dir)
    
    if not pdf_files:
        print("\n未找到任何 PDF 文件")
        return
    
    # 获取数据库信息
    db_files, db_folders = get_database_pdfs(zotero_data_dir)
    
    print(f"\n统计:")
    print(f"  - 总 PDF 文件: {len(pdf_files)}")
    print(f"  - 数据库记录: {len(db_files)}")
    print(f"  - 有效文件夹: {len(db_folders)}")
    
    # 步骤1: 清理重复的 PDF
    print(f"\n{'='*60}")
    print("步骤 1: 清理重复的 PDF 文件")
    print(f"{'='*60}")
    dup_count = clean_duplicate_pdfs(pdf_files, db_files, back_dir)
    
    # 重新扫描（因为删除了一些文件）
    if dup_count > 0:
        pdf_files = collect_pdf_files(storage_dir)
    
    # 步骤2: 清理孤立的 PDF
    print(f"\n{'='*60}")
    print("步骤 2: 清理孤立的 PDF 文件")
    print(f"{'='*60}")
    orphan_count = clean_orphaned_pdfs(pdf_files, db_files, back_dir)
    
    # 步骤3: 清理空文件夹和无效文件夹
    print(f"\n{'='*60}")
    print("步骤 3: 清理空文件夹和无效文件夹")
    print(f"{'='*60}")
    folder_count = clean_empty_folders(storage_dir, db_folders)
    
    # 最终统计
    print(f"\n{'='*60}")
    print("清理完成！")
    print(f"{'='*60}")
    print(f"  - 删除重复文件: {dup_count} 个")
    print(f"  - 删除孤立文件: {orphan_count} 个")
    print(f"  - 删除文件夹: {folder_count} 个")
    print(f"  - 备份位置: {back_dir}")
    print(f"{'='*60}")


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n用户中断，程序退出")
        sys.exit(0)
    except Exception as e:
        print(f"\n发生错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
