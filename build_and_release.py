# -*- coding: utf-8 -*-
"""
构建和发布脚本 - Y2订单处理辅助工具
功能：
1. 构建 PyInstaller 打包
2. 创建版本压缩包
3. 计算文件哈希
4. 生成 version.json
5. 输出发布文件到 release 目录
"""

import os
import sys
import json
import hashlib
import shutil
import subprocess
import zipfile
from pathlib import Path
from datetime import datetime

# 配置
APP_NAME = "Y2订单处理辅助工具"
VERSION = "1.9.0"
SPEC_FILE = "Y2订单处理辅助工具1.9.spec"
DIST_DIR = "dist"
BUILD_DIR = "build"
RELEASE_DIR = "release"


def log(message):
    """打印日志"""
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")


def clean_build():
    """清理构建目录"""
    log("清理构建目录...")
    
    dirs_to_clean = [BUILD_DIR, DIST_DIR, RELEASE_DIR]
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            log(f"  已删除: {dir_name}")
    
    log("清理完成")


def build_app():
    """使用 PyInstaller 构建应用"""
    log("开始构建应用...")
    
    if not os.path.exists(SPEC_FILE):
        log(f"错误: 找不到 spec 文件 {SPEC_FILE}")
        return False
    
    try:
        result = subprocess.run(
            [sys.executable, "-m", "PyInstaller", SPEC_FILE, "--clean"],
            capture_output=True,
            text=True,
            check=True
        )
        log("构建成功")
        return True
    except subprocess.CalledProcessError as e:
        log(f"构建失败: {e}")
        log(f"错误输出: {e.stderr}")
        return False


def calculate_hash(file_path):
    """计算文件 SHA256 哈希"""
    sha256 = hashlib.sha256()
    with open(file_path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            sha256.update(chunk)
    return f"sha256:{sha256.hexdigest()}"


def create_release_package():
    """创建发布压缩包"""
    log("创建发布压缩包...")
    
    # 确保发布目录存在
    os.makedirs(RELEASE_DIR, exist_ok=True)
    
    # 源目录（PyInstaller 输出）
    source_dir = os.path.join(DIST_DIR, f"{APP_NAME}{VERSION}")
    if not os.path.exists(source_dir):
        # 尝试其他可能的目录名
        possible_names = [
            f"{APP_NAME}1.9",
            APP_NAME,
            "Y2订单处理辅助工具1.9"
        ]
        for name in possible_names:
            test_dir = os.path.join(DIST_DIR, name)
            if os.path.exists(test_dir):
                source_dir = test_dir
                break
    
    if not os.path.exists(source_dir):
        log(f"错误: 找不到构建输出目录")
        return None
    
    # 创建压缩包
    zip_name = f"{APP_NAME}{VERSION}.zip"
    zip_path = os.path.join(RELEASE_DIR, zip_name)
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, source_dir)
                zipf.write(file_path, arcname)
    
    file_size = os.path.getsize(zip_path)
    log(f"压缩包已创建: {zip_path}")
    log(f"  大小: {file_size / 1024 / 1024:.2f} MB")
    
    return zip_path


def generate_version_json(zip_path, changelog=None):
    """生成 version.json"""
    log("生成 version.json...")
    
    if changelog is None:
        changelog = [
            "新增远程自动更新功能",
            "优化图片处理性能",
            "修复已知问题"
        ]
    
    version_info = {
        "version": VERSION,
        "min_version": "1.8.0",
        "download_url": f"https://github.com/yourname/Y2Tool/releases/download/v{VERSION}/{APP_NAME}{VERSION}.zip",
        "changelog": changelog,
        "force_update": False,
        "file_size": os.path.getsize(zip_path),
        "hash": calculate_hash(zip_path),
        "release_date": datetime.now().strftime("%Y-%m-%d")
    }
    
    version_path = os.path.join(RELEASE_DIR, "version.json")
    with open(version_path, 'w', encoding='utf-8') as f:
        json.dump(version_info, f, ensure_ascii=False, indent=2)
    
    log(f"version.json 已生成: {version_path}")
    return version_path


def copy_installer_files():
    """复制安装程序相关文件"""
    log("复制安装程序文件...")
    
    files_to_copy = [
        "installer_script.iss",
        "logo.ico",
        "make_installer.bat"
    ]
    
    for file in files_to_copy:
        if os.path.exists(file):
            shutil.copy2(file, RELEASE_DIR)
            log(f"  已复制: {file}")


def print_release_notes():
    """打印发布说明"""
    print("\n" + "=" * 60)
    print("发布完成!")
    print("=" * 60)
    print(f"\n版本: {VERSION}")
    print(f"发布目录: {os.path.abspath(RELEASE_DIR)}")
    print("\n发布文件:")
    
    for file in os.listdir(RELEASE_DIR):
        file_path = os.path.join(RELEASE_DIR, file)
        size = os.path.getsize(file_path)
        if size > 1024 * 1024:
            size_str = f"{size / 1024 / 1024:.2f} MB"
        else:
            size_str = f"{size / 1024:.2f} KB"
        print(f"  - {file} ({size_str})")
    
    print("\n下一步操作:")
    print("  1. 上传 release 目录中的文件到 GitHub Releases")
    print("  2. 更新 version.json 中的 download_url 为实际地址")
    print("  3. 运行 make_installer.bat 创建安装程序（可选）")
    print("=" * 60)


def main():
    """主函数"""
    print("=" * 60)
    print(f"Y2订单处理辅助工具 - 构建发布脚本 v{VERSION}")
    print("=" * 60)
    
    # 检查命令行参数
    if len(sys.argv) > 1:
        if sys.argv[1] == "--clean":
            clean_build()
            return
    
    # 完整构建流程
    clean_build()
    
    if not build_app():
        log("构建失败，退出")
        sys.exit(1)
    
    zip_path = create_release_package()
    if not zip_path:
        log("创建压缩包失败，退出")
        sys.exit(1)
    
    generate_version_json(zip_path)
    copy_installer_files()
    
    print_release_notes()


if __name__ == '__main__':
    main()
