"""
MTP Backup Tool
===============
将本地文件夹内容通过 MTP 协议备份到中兴 F50 内部共享存储空间。

核心原理：
  - 使用 Windows Shell.Application COM 接口操作 MTP 设备
  - 通过 Shell Namespace 遍历设备目录树
  - 使用 Folder.CopyHere() 发起文件传输，轮询等待完成
  - 支持中文文件名（通过 ParseName 获取 FolderItem 对象）

依赖：
  - pywin32 (win32com.client)
  - Windows 10/11，F50 通过 USB 以 MTP 模式连接

用法：
  python mtp_backup.py
"""

import os
import sys
import time
import logging
import win32com.client
import pythoncom
from datetime import datetime
from pathlib import Path
from win32com.shell import shellcon

# ─────────────────────────────────────────────────────────────────────────────
# 配置区（按需修改）
# ─────────────────────────────────────────────────────────────────────────────
CONFIG = {
    # 本地源文件夹路径列表（支持多个目录）
    "source_dirs": [
        r"F:\share\sync",
    ],

    # F50 设备名称（模糊匹配，大小写不敏感）
    "device_name": "F50",

    # F50 内部存储上的备份根目录名称
    "backup_root_name": "PC备份",

    # 备份子目录的日期格式（strftime 格式）
    "date_format": "%Y%m%d",

    # 单文件传输超时时间（秒）
    "copy_timeout_sec": 60,

    # 传输完成轮询间隔（秒）
    "poll_interval_sec": 2,

    # 日志文件路径
    "log_file": r"C:\Users\Administrator\mtp-backup\backup.log",
}
# ─────────────────────────────────────────────────────────────────────────────


def setup_logging(log_file: str) -> logging.Logger:
    """
    配置日志系统：同时输出到控制台和文件。

    Args:
        log_file: 日志文件路径

    Returns:
        配置好的 Logger 实例
    """
    logger = logging.getLogger("mtp_backup")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter(
        "[%(asctime)s] %(levelname)-5s - %(message)s",
        "%Y-%m-%d %H:%M:%S"
    )

    # 控制台 Handler
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    # 文件 Handler（UTF-8 编码，支持中文）
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger


# ─────────────────────────────────────────────────────────────────────────────
# Shell Namespace / MTP 操作辅助函数
# ─────────────────────────────────────────────────────────────────────────────

def get_f50_storage_folder(shell, device_name: str, logger: logging.Logger):
    """
    在"此电脑"下定位 F50 设备的内部共享存储空间。

    遍历 Shell Namespace 的 CSIDL_DRIVES（0x11，即"此电脑"），
    找到名称匹配 device_name 的便携设备，取其第一个存储空间。

    Args:
        shell:       Shell.Application COM 对象
        device_name: 设备名称关键字（模糊匹配）
        logger:      日志实例

    Returns:
        存储空间的 Folder COM 对象，未找到时返回 None
    """
    computer = shell.Namespace(shellcon.CSIDL_DRIVES)  # 0x11 = 此电脑
    items = computer.Items()

    for i in range(items.Count):
        item = items.Item(i)
        if device_name.lower() in item.Name.lower():
            logger.info(f"找到设备: {item.Name}")
            dev_ns = shell.Namespace(item.Path)
            if dev_ns is None:
                logger.error("无法打开设备命名空间")
                return None

            storage_items = dev_ns.Items()
            if storage_items.Count == 0:
                logger.error("设备下无存储空间")
                return None

            # 取第一个存储（内部共享存储空间）
            storage = storage_items.Item(0)
            logger.info(f"存储空间: {storage.Name}")
            return storage.GetFolder

    logger.error(f"未找到设备: {device_name}，请确认 F50 已通过 USB 连接")
    return None


def find_subfolder(parent_folder, name: str):
    """
    在 MTP Folder 下查找指定名称的子文件夹。

    Args:
        parent_folder: 父级 Folder COM 对象
        name:          要查找的文件夹名称（精确匹配）

    Returns:
        匹配的 FolderItem COM 对象，未找到返回 None
    """
    items = parent_folder.Items()
    for i in range(items.Count):
        item = items.Item(i)
        if item.Name == name and item.IsFolder:
            return item
    return None


def ensure_mtp_folder(parent_folder, name: str, logger: logging.Logger, wait_sec: int = 10):
    """
    确保 MTP 目录存在：已存在则直接返回，不存在则创建并等待。

    MTP 设备的 NewFolder 是异步操作，需要轮询等待目录真正出现。

    Args:
        parent_folder: 父级 Folder COM 对象
        name:          目录名称
        logger:        日志实例
        wait_sec:      等待创建完成的最大秒数

    Returns:
        子目录的 Folder COM 对象，失败返回 None
    """
    existing = find_subfolder(parent_folder, name)
    if existing:
        logger.debug(f"已存在目录: {name}")
        return existing.GetFolder

    logger.info(f"创建目录: {name}")
    parent_folder.NewFolder(name)

    # 轮询等待目录出现
    deadline = time.time() + wait_sec
    while time.time() < deadline:
        time.sleep(0.5)
        found = find_subfolder(parent_folder, name)
        if found:
            logger.debug(f"目录创建成功: {name}")
            return found.GetFolder

    logger.error(f"创建目录超时: {name}")
    return None


def get_mtp_filenames(mtp_folder) -> set:
    """
    获取 MTP 文件夹下所有文件/子目录的名称集合。

    用于判断文件是否已存在（增量备份跳过逻辑）。

    Args:
        mtp_folder: MTP Folder COM 对象

    Returns:
        文件名集合（set of str）
    """
    if mtp_folder is None:
        return set()
    items = mtp_folder.Items()
    return {items.Item(i).Name for i in range(items.Count)}


def copy_file_to_mtp(
    shell,
    local_path: str,
    mtp_folder,
    filename: str,
    logger: logging.Logger,
    timeout: int = 60
) -> bool:
    """
    将单个本地文件复制到 MTP 目标目录，并等待传输完成。

    流程：
      1. 通过 Shell.Namespace 获取源目录的命名空间
      2. 用 ParseName 获取文件的 FolderItem 对象（支持中文文件名）
      3. 调用 Folder.CopyHere 发起异步传输（flag 0x14 = 静默模式）
      4. 轮询目标目录，直到文件名出现或超时

    Args:
        shell:      Shell.Application COM 对象
        local_path: 本地文件完整路径
        mtp_folder: 目标 MTP Folder COM 对象
        filename:   文件名（不含路径）
        logger:     日志实例
        timeout:    等待超时秒数

    Returns:
        True 表示传输成功，False 表示失败或超时
    """
    try:
        src_dir = os.path.dirname(local_path)
        src_ns = shell.Namespace(src_dir)
        if src_ns is None:
            logger.error(f"无法打开源目录: {src_dir}")
            return False

        # ParseName 支持中文文件名，比直接拼路径更可靠
        file_item = src_ns.ParseName(filename)
        if file_item is None:
            logger.error(f"无法获取文件对象: {filename}")
            return False

        # 0x14 = FOF_SILENT(0x4) | FOF_NOCONFIRMATION(0x10)，静默复制不弹窗
        mtp_folder.CopyHere(file_item, 0x14)

        # 轮询等待文件出现在目标目录
        deadline = time.time() + timeout
        while time.time() < deadline:
            time.sleep(CONFIG["poll_interval_sec"])
            if filename in get_mtp_filenames(mtp_folder):
                return True

        logger.error(f"传输超时 ({timeout}s): {filename}")
        return False

    except Exception as e:
        logger.error(f"传输异常 [{filename}]: {e}")
        return False


# ─────────────────────────────────────────────────────────────────────────────
# 备份核心逻辑
# ─────────────────────────────────────────────────────────────────────────────

def backup_directory(
    shell,
    source_dir: str,
    mtp_target_folder,
    logger: logging.Logger
) -> dict:
    """
    递归备份本地目录到 MTP 目标文件夹。

    - 保留原始目录结构（相对路径）
    - 已存在的文件自动跳过（增量备份）
    - 逐层创建 MTP 子目录

    Args:
        shell:             Shell.Application COM 对象
        source_dir:        本地源目录路径
        mtp_target_folder: MTP 目标 Folder COM 对象
        logger:            日志实例

    Returns:
        统计字典 {"copied": int, "skipped": int, "failed": int}
    """
    stats = {"copied": 0, "skipped": 0, "failed": 0}
    source_path = Path(source_dir)

    if not source_path.exists():
        logger.error(f"源目录不存在: {source_dir}")
        return stats

    # 收集所有文件（递归）
    all_files = [f for f in source_path.rglob("*") if f.is_file()]
    logger.info(f"发现 {len(all_files)} 个文件")

    for file_path in all_files:
        rel_path = file_path.relative_to(source_path)
        rel_parts = rel_path.parts  # 相对路径各级目录名

        # 逐层确保 MTP 目录结构存在
        current_folder = mtp_target_folder
        dir_ok = True
        for part in rel_parts[:-1]:  # 最后一个元素是文件名，跳过
            current_folder = ensure_mtp_folder(current_folder, part, logger)
            if current_folder is None:
                logger.error(f"无法创建子目录: {part}，跳过文件: {rel_path}")
                stats["failed"] += 1
                dir_ok = False
                break

        if not dir_ok:
            continue

        fname = file_path.name

        # 增量备份：已存在则跳过
        if fname in get_mtp_filenames(current_folder):
            logger.info(f"[SKIP] {rel_path}")
            stats["skipped"] += 1
            continue

        # 执行文件传输
        logger.info(f"[COPY] {rel_path}")
        success = copy_file_to_mtp(
            shell, str(file_path), current_folder, fname, logger,
            timeout=CONFIG["copy_timeout_sec"]
        )

        if success:
            logger.info(f"[OK]   {fname}")
            stats["copied"] += 1
        else:
            stats["failed"] += 1

    return stats


# ─────────────────────────────────────────────────────────────────────────────
# 主入口
# ─────────────────────────────────────────────────────────────────────────────

def run_backup():
    """
    备份主流程：
      1. 初始化 COM 环境和日志
      2. 定位 F50 内部存储空间
      3. 创建 备份根目录 / 日期目录 / 源目录名 三级目录结构
      4. 逐个备份 source_dirs 中的目录
      5. 输出汇总统计
    """
    # Windows COM 必须在使用前初始化
    pythoncom.CoInitialize()
    logger = setup_logging(CONFIG["log_file"])

    logger.info("=" * 60)
    logger.info("MTP 备份开始")
    logger.info("=" * 60)

    try:
        shell = win32com.client.Dispatch("Shell.Application")

        # ── Step 1: 定位 F50 存储 ──────────────────────────────
        storage_folder = get_f50_storage_folder(shell, CONFIG["device_name"], logger)
        if storage_folder is None:
            logger.error("无法访问 F50 存储空间，请确认设备已通过 USB 连接")
            sys.exit(1)

        # ── Step 2: 创建备份目录结构 ───────────────────────────
        # 结构：<backup_root_name>/<YYYYMMDD>/<源目录名>/
        backup_root = ensure_mtp_folder(
            storage_folder, CONFIG["backup_root_name"], logger
        )
        if backup_root is None:
            logger.error("无法创建备份根目录")
            sys.exit(1)

        date_str = datetime.now().strftime(CONFIG["date_format"])
        date_folder = ensure_mtp_folder(backup_root, date_str, logger)
        if date_folder is None:
            logger.error("无法创建日期目录")
            sys.exit(1)

        logger.info(
            f"备份目标: F50\\内部共享存储空间\\"
            f"{CONFIG['backup_root_name']}\\{date_str}"
        )

        # ── Step 3: 逐目录备份 ────────────────────────────────
        total_stats = {"copied": 0, "skipped": 0, "failed": 0}

        for source_dir in CONFIG["source_dirs"]:
            # 在日期目录下以源目录名创建子目录
            dir_name = os.path.basename(source_dir.rstrip("\\/"))
            dir_folder = ensure_mtp_folder(date_folder, dir_name, logger)
            if dir_folder is None:
                logger.error(f"无法为 [{dir_name}] 创建目标目录，跳过")
                continue

            stats = backup_directory(shell, source_dir, dir_folder, logger)
            for k in total_stats:
                total_stats[k] += stats[k]

        # ── Step 4: 汇总 ──────────────────────────────────────
        logger.info("=" * 60)
        logger.info(
            f"备份完成！"
            f"[OK] {total_stats['copied']}  "
            f"[SKIP] {total_stats['skipped']}  "
            f"[FAIL] {total_stats['failed']}"
        )
        logger.info("=" * 60)

    except Exception as e:
        logger.exception(f"备份过程发生未预期异常: {e}")
        sys.exit(1)
    finally:
        # 释放 COM 资源
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    run_backup()
