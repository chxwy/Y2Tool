# -*- coding: utf-8 -*-
"""
更新助手程序 - 独立的更新执行器
此程序在后台运行，负责：
1. 等待主程序关闭
2. 解压更新包
3. 替换旧文件
4. 重启主程序
5. 自删除
"""

import os
import sys
import time
import shutil
import zipfile
import subprocess
import tempfile
from pathlib import Path


def log(message):
    """记录日志"""
    log_path = os.path.join(tempfile.gettempdir(), 'Y2_updater.log')
    with open(log_path, 'a', encoding='utf-8') as f:
        f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")
    print(message)


def wait_for_process_exit(process_path, timeout=30):
    """等待进程退出"""
    import psutil
    
    process_name = os.path.basename(process_path)
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        found = False
        for proc in psutil.process_iter(['name', 'exe']):
            try:
                if proc.info['name'] == process_name or proc.info['exe'] == process_path:
                    found = True
                    break
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        
        if not found:
            log(f"原程序已退出: {process_name}")
            return True
        
        time.sleep(0.5)
    
    log(f"等待原程序退出超时")
    return False


def kill_process(process_path):
    """强制结束进程"""
    import psutil
    
    process_name = os.path.basename(process_path)
    killed = False
    
    for proc in psutil.process_iter(['name', 'exe']):
        try:
            if proc.info['name'] == process_name or proc.info['exe'] == process_path:
                proc.kill()
                killed = True
                log(f"已结束进程: {process_name}")
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    
    return killed


def backup_old_version(target_dir):
    """备份旧版本"""
    backup_dir = os.path.join(tempfile.gettempdir(), 'Y2_backup', time.strftime('%Y%m%d_%H%M%S'))
    
    try:
        os.makedirs(backup_dir, exist_ok=True)
        
        # 备份主程序
        exe_path = os.path.join(target_dir, 'Y2订单处理辅助工具.exe')
        if os.path.exists(exe_path):
            shutil.copy2(exe_path, backup_dir)
            log(f"已备份主程序到: {backup_dir}")
        
        # 备份配置文件
        config_files = ['config.json', 'processing_config.json']
        for config_file in config_files:
            config_path = os.path.join(target_dir, config_file)
            if os.path.exists(config_path):
                shutil.copy2(config_path, backup_dir)
        
        return backup_dir
    except Exception as e:
        log(f"备份失败: {e}")
        return None


def extract_update(zip_path, target_dir):
    """解压更新包"""
    try:
        log(f"开始解压: {zip_path}")
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # 获取压缩包内的根目录名
            root_dirs = set()
            for name in zip_ref.namelist():
                parts = name.split('/')
                if len(parts) > 1:
                    root_dirs.add(parts[0])
            
            # 解压到临时目录
            temp_extract_dir = os.path.join(tempfile.gettempdir(), 'Y2_update_extract')
            if os.path.exists(temp_extract_dir):
                shutil.rmtree(temp_extract_dir)
            os.makedirs(temp_extract_dir)
            
            zip_ref.extractall(temp_extract_dir)
            log(f"解压完成到临时目录: {temp_extract_dir}")
            
            return temp_extract_dir, root_dirs
            
    except Exception as e:
        log(f"解压失败: {e}")
        return None, None


def replace_files(source_dir, target_dir, root_dirs):
    """替换文件"""
    try:
        log(f"开始替换文件: {source_dir} -> {target_dir}")
        
        # 如果压缩包内有根目录，进入该目录
        if len(root_dirs) == 1:
            source_dir = os.path.join(source_dir, list(root_dirs)[0])
        
        # 遍历并替换文件
        for root, dirs, files in os.walk(source_dir):
            # 计算相对路径
            rel_path = os.path.relpath(root, source_dir)
            target_path = os.path.join(target_dir, rel_path) if rel_path != '.' else target_dir
            
            # 创建目标目录
            os.makedirs(target_path, exist_ok=True)
            
            # 复制文件
            for file in files:
                src_file = os.path.join(root, file)
                dst_file = os.path.join(target_path, file)
                
                # 跳过正在运行的更新助手
                if file.lower() in ['updater.exe', 'updater.py']:
                    continue
                
                # 如果目标文件存在且正在使用，重命名旧文件
                if os.path.exists(dst_file):
                    try:
                        os.remove(dst_file)
                    except:
                        backup_name = f"{dst_file}.old"
                        if os.path.exists(backup_name):
                            os.remove(backup_name)
                        os.rename(dst_file, backup_name)
                
                shutil.copy2(src_file, dst_file)
                log(f"已替换: {dst_file}")
        
        log("文件替换完成")
        return True
        
    except Exception as e:
        log(f"替换文件失败: {e}")
        return False


def restart_application(exe_path):
    """重启应用程序"""
    try:
        log(f"重启程序: {exe_path}")
        
        if os.path.exists(exe_path):
            subprocess.Popen([exe_path], shell=False)
            log("程序已重启")
            return True
        else:
            log(f"程序文件不存在: {exe_path}")
            return False
            
    except Exception as e:
        log(f"重启失败: {e}")
        return False


def cleanup(zip_path, extract_dir):
    """清理临时文件"""
    try:
        if os.path.exists(zip_path):
            os.remove(zip_path)
            log(f"已删除更新包: {zip_path}")
        
        if extract_dir and os.path.exists(extract_dir):
            shutil.rmtree(extract_dir)
            log(f"已清理解压目录: {extract_dir}")
            
    except Exception as e:
        log(f"清理失败: {e}")


def self_delete():
    """自删除"""
    try:
        # 创建一个批处理文件来删除自己
        batch_path = os.path.join(tempfile.gettempdir(), 'delete_updater.bat')
        updater_path = sys.executable if getattr(sys, 'frozen', False) else sys.argv[0]
        
        with open(batch_path, 'w', encoding='gbk') as f:
            f.write('@echo off\n')
            f.write('timeout /t 2 /nobreak >nul\n')
            f.write(f'del "{updater_path}"\n')
            f.write(f'del "%~f0"\n')
        
        subprocess.Popen([batch_path], shell=True)
        log("已启动自删除脚本")
        
    except Exception as e:
        log(f"自删除失败: {e}")


def main():
    """主函数"""
    if len(sys.argv) < 3:
        log("用法: updater.py <update_zip_path> <target_dir> [main_exe_path]")
        sys.exit(1)
    
    update_zip_path = sys.argv[1]
    target_dir = sys.argv[2]
    main_exe_path = sys.argv[3] if len(sys.argv) > 3 else None
    
    log("=" * 50)
    log("更新助手启动")
    log(f"更新包: {update_zip_path}")
    log(f"目标目录: {target_dir}")
    log(f"主程序: {main_exe_path}")
    
    # 1. 等待主程序退出
    if main_exe_path and os.path.exists(main_exe_path):
        log("等待主程序退出...")
        if not wait_for_process_exit(main_exe_path, timeout=10):
            log("主程序未正常退出，尝试强制结束...")
            kill_process(main_exe_path)
            time.sleep(1)
    
    # 2. 备份旧版本
    backup_dir = backup_old_version(target_dir)
    
    # 3. 解压更新包
    extract_dir, root_dirs = extract_update(update_zip_path, target_dir)
    if not extract_dir:
        log("解压失败，更新终止")
        if backup_dir:
            log(f"可以从备份恢复: {backup_dir}")
        sys.exit(1)
    
    # 4. 替换文件
    if not replace_files(extract_dir, target_dir, root_dirs):
        log("文件替换失败")
        sys.exit(1)
    
    # 5. 清理临时文件
    cleanup(update_zip_path, extract_dir)
    
    # 6. 重启主程序
    if main_exe_path:
        restart_application(main_exe_path)
    else:
        # 尝试找到主程序
        possible_exe = os.path.join(target_dir, 'Y2订单处理辅助工具.exe')
        if os.path.exists(possible_exe):
            restart_application(possible_exe)
    
    log("更新完成")
    log("=" * 50)
    
    # 7. 自删除
    self_delete()


if __name__ == '__main__':
    main()
