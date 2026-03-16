"""
MTP Watch & Backup
==================
监听本地目录的文件变化，当有新文件写入完成时，
自动将该文件通过 MTP 协议备份到 F50 内部存储。

监听原理：
  - 使用 watchdog 库监听文件系统事件
  - 监听 on_created / on_moved 事件（覆盖复制粘贴场景）
  - 文件写入完成检测：连续 2 次轮询文件大小不变则认为写入完成
  - 防抖：同一文件 3 秒内重复事件只处理一次

备份原理：
  - 复用 mtp_backup.py 中的 Shell COM 传输逻辑
  - 保留相对路径，在 F50 上重建目录结构
  - 目标路径：F50\内部共享存储空间\PC备份\YYYYMMDD\<源目录名>\...

用法：
  python mtp_watch.py              # 前台运行，Ctrl+C 停止
  python mtp_watch.py --daemon     # 后台运行（写 PID 文件）
"""

import os
import sys
import time
import logging
import argparse
import threading
from datetime import datetime
from pathlib import Path

import win32com.client
import pythoncom
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from win32com.shell import shellcon

# ─────────────────────────────────────────────────────────────────────────────
# 配置区（按需修改）
# ─────────────────────────────────────────────────────────────────────────────
CONFIG = {
    # 监听的本地目录（支持多个）
    "watch_dirs": [
        r"F:\share\sync",
    ],

    # F50 设备名称（模糊匹配）
    "device_name": "F50",

    # F50 上的备份根目录名
    "backup_root_name": "PC备份",

    # 日期格式（备份子目录命名）
    "date_format": "%Y%m%d",

    # 文件写入完成检测：连续 N 次大小不变则认为写入完成
    "stable_checks": 3,

    # 写入完成检测间隔（秒）
    "stable_interval": 1.0,

    # 防抖窗口（秒）：同一文件在此时间内重复触发只处理一次
    "debounce_sec": 3.0,

    # 单文件 MTP 传输超时（秒）
    "copy_timeout_sec": 60,

    # MTP 传输完成轮询间隔（秒）
    "poll_interval_sec": 2,

    # 忽略的文件扩展名（临时文件）
    "ignore_extensions": {".tmp", ".part", ".crdownload", ".download", "~"},

    # 忽略的文件名前缀
    "ignore_prefixes": {"~$", "."},

    # 日志文件路径
    "log_file": r"C:\Users\Administrator\mtp-backup\watch.log",

    # PID 文件路径（daemon 模式使用）
    "pid_file": r"C:\Users\Administrator\mtp-backup\watch.pid",
}
# ─────────────────────────────────────────────────────────────────────────────


def setup_logging(log_file: str) -> logging.Logger:
    """配置日志：控制台 + 文件双输出，UTF-8 编码"""
    logger = logging.getLogger("mtp_watch")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter(
        "[%(asctime)s] %(levelname)-5s - %(message)s",
        "%Y-%m-%d %H:%M:%S"
    )
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    return logger


# ─────────────────────────────────────────────────────────────────────────────
# MTP 操作（复用 mtp_backup.py 逻辑）
# ─────────────────────────────────────────────────────────────────────────────

def get_storage_folder(shell, device_name: str, logger):
    """定位 F50 内部存储的 Folder COM 对象"""
    computer = shell.Namespace(shellcon.CSIDL_DRIVES)
    items = computer.Items()
    for i in range(items.Count):
        item = items.Item(i)
        if device_name.lower() in item.Name.lower():
            dev_ns = shell.Namespace(item.Path)
            if dev_ns and dev_ns.Items().Count > 0:
                storage = dev_ns.Items().Item(0)
                logger.debug(f"存储空间: {storage.Name}")
                return storage.GetFolder
    logger.error(f"未找到设备: {device_name}")
    return None


def find_subfolder(parent_folder, name: str):
    """在 MTP Folder 下查找子文件夹"""
    items = parent_folder.Items()
    for i in range(items.Count):
        item = items.Item(i)
        if item.Name == name and item.IsFolder:
            return item
    return None


def ensure_mtp_folder(parent_folder, name: str, logger, wait_sec: int = 10):
    """确保 MTP 子目录存在，不存在则创建并等待"""
    existing = find_subfolder(parent_folder, name)
    if existing:
        return existing.GetFolder

    logger.info(f"  创建目录: {name}")
    parent_folder.NewFolder(name)

    deadline = time.time() + wait_sec
    while time.time() < deadline:
        time.sleep(0.5)
        found = find_subfolder(parent_folder, name)
        if found:
            return found.GetFolder

    logger.error(f"  创建目录超时: {name}")
    return None


def get_mtp_filenames(mtp_folder) -> set:
    """获取 MTP 目录下所有文件名"""
    if mtp_folder is None:
        return set()
    items = mtp_folder.Items()
    return {items.Item(i).Name for i in range(items.Count)}


def copy_file_to_mtp(shell, local_path: str, mtp_folder, filename: str,
                     logger, timeout: int = 60) -> bool:
    """
    将单个文件复制到 MTP 目录并等待完成。
    使用 ParseName 获取 FolderItem，支持中文文件名。
    """
    try:
        src_dir = os.path.dirname(local_path)
        src_ns = shell.Namespace(src_dir)
        if src_ns is None:
            logger.error(f"  无法打开目录: {src_dir}")
            return False

        file_item = src_ns.ParseName(filename)
        if file_item is None:
            logger.error(f"  无法获取文件对象: {filename}")
            return False

        # 0x14 = FOF_SILENT | FOF_NOCONFIRMATION，静默复制
        mtp_folder.CopyHere(file_item, 0x14)

        # 轮询等待文件出现
        deadline = time.time() + timeout
        while time.time() < deadline:
            time.sleep(CONFIG["poll_interval_sec"])
            if filename in get_mtp_filenames(mtp_folder):
                return True

        logger.error(f"  传输超时 ({timeout}s): {filename}")
        return False

    except Exception as e:
        logger.error(f"  传输异常 [{filename}]: {e}")
        return False


def backup_single_file(local_path: str, watch_root: str, logger) -> bool:
    """
    备份单个文件到 F50。

    目标路径结构：
      F50\内部共享存储空间\<backup_root>\<YYYYMMDD>\<watch_root_name>\<相对路径>

    Args:
        local_path:  本地文件完整路径
        watch_root:  监听的根目录（用于计算相对路径）
        logger:      日志实例

    Returns:
        True 表示成功
    """
    # 每次备份都重新初始化 COM（线程安全）
    pythoncom.CoInitialize()
    try:
        shell = win32com.client.Dispatch("Shell.Application")

        # 1. 定位存储
        storage_folder = get_storage_folder(shell, CONFIG["device_name"], logger)
        if storage_folder is None:
            return False

        # 2. 构建目标目录：backup_root / date / watch_dir_name / 相对子路径
        backup_root = ensure_mtp_folder(
            storage_folder, CONFIG["backup_root_name"], logger
        )
        if backup_root is None:
            return False

        date_str = datetime.now().strftime(CONFIG["date_format"])
        date_folder = ensure_mtp_folder(backup_root, date_str, logger)
        if date_folder is None:
            return False

        watch_dir_name = os.path.basename(watch_root.rstrip("\\/"))
        dir_folder = ensure_mtp_folder(date_folder, watch_dir_name, logger)
        if dir_folder is None:
            return False

        # 3. 计算相对路径，逐层创建子目录
        rel_path = Path(local_path).relative_to(watch_root)
        rel_parts = rel_path.parts  # (子目录..., 文件名)

        current_folder = dir_folder
        for part in rel_parts[:-1]:
            current_folder = ensure_mtp_folder(current_folder, part, logger)
            if current_folder is None:
                return False

        filename = rel_parts[-1]

        # 4. 增量检查：已存在则跳过
        if filename in get_mtp_filenames(current_folder):
            logger.info(f"  [SKIP] 已存在: {rel_path}")
            return True

        # 5. 执行传输
        logger.info(f"  [COPY] {rel_path}")
        ok = copy_file_to_mtp(
            shell, local_path, current_folder, filename, logger,
            timeout=CONFIG["copy_timeout_sec"]
        )
        if ok:
            logger.info(
                f"  [OK]   {filename} -> "
                f"F50\\{CONFIG['backup_root_name']}\\{date_str}\\"
                f"{watch_dir_name}\\{rel_path}"
            )
        return ok

    finally:
        pythoncom.CoUninitialize()


# ─────────────────────────────────────────────────────────────────────────────
# 文件写入完成检测
# ─────────────────────────────────────────────────────────────────────────────

def wait_for_file_stable(path: str, logger) -> bool:
    """
    等待文件写入完成：连续 stable_checks 次检测文件大小不变则认为完成。

    处理大文件复制粘贴时文件仍在写入的情况。

    Args:
        path:   文件路径
        logger: 日志实例

    Returns:
        True 表示文件稳定可读，False 表示超时或文件消失
    """
    stable_checks = CONFIG["stable_checks"]
    interval = CONFIG["stable_interval"]
    max_wait = 120  # 最多等待 120 秒

    last_size = -1
    stable_count = 0
    waited = 0

    while waited < max_wait:
        try:
            if not os.path.exists(path):
                return False
            size = os.path.getsize(path)
            if size == last_size:
                stable_count += 1
                if stable_count >= stable_checks:
                    return True
            else:
                stable_count = 0
                last_size = size
        except OSError:
            pass

        time.sleep(interval)
        waited += interval

    logger.warning(f"等待文件稳定超时: {path}")
    return False


# ─────────────────────────────────────────────────────────────────────────────
# watchdog 事件处理器
# ─────────────────────────────────────────────────────────────────────────────

class MTPBackupHandler(FileSystemEventHandler):
    """
    监听文件系统事件，触发 MTP 备份。

    处理的事件：
      - on_created：新文件创建（复制粘贴、下载完成等）
      - on_moved：文件移动/重命名到监听目录（部分软件先写临时文件再重命名）

    防抖机制：
      - 用 _recent 字典记录最近处理的文件路径和时间戳
      - debounce_sec 内重复触发的同一文件只处理一次
    """

    def __init__(self, watch_root: str, logger: logging.Logger):
        super().__init__()
        self.watch_root = watch_root
        self.logger = logger
        self._recent: dict[str, float] = {}  # path -> last_trigger_time
        self._lock = threading.Lock()

    def _should_ignore(self, path: str) -> bool:
        """判断文件是否应该忽略（临时文件、隐藏文件等）"""
        name = os.path.basename(path)
        ext = Path(path).suffix.lower()

        # 忽略特定扩展名
        if ext in CONFIG["ignore_extensions"]:
            return True

        # 忽略特定前缀
        for prefix in CONFIG["ignore_prefixes"]:
            if name.startswith(prefix):
                return True

        # 只处理文件，不处理目录
        if os.path.isdir(path):
            return True

        return False

    def _debounce(self, path: str) -> bool:
        """
        防抖检查：返回 True 表示应该处理，False 表示在防抖窗口内重复触发。
        """
        now = time.time()
        with self._lock:
            last = self._recent.get(path, 0)
            if now - last < CONFIG["debounce_sec"]:
                return False
            self._recent[path] = now
            # 清理过期记录，避免内存泄漏
            expired = [k for k, v in self._recent.items()
                       if now - v > CONFIG["debounce_sec"] * 10]
            for k in expired:
                del self._recent[k]
        return True

    def _handle_file(self, path: str, event_type: str):
        """
        处理单个文件事件的核心逻辑：
          1. 过滤检查
          2. 防抖检查
          3. 等待文件写入完成
          4. 触发 MTP 备份
        """
        if self._should_ignore(path):
            return

        if not self._debounce(path):
            self.logger.debug(f"[防抖] 跳过重复事件: {os.path.basename(path)}")
            return

        self.logger.info(f"[{event_type}] 检测到新文件: {path}")

        # 等待文件写入完成（处理大文件复制中的情况）
        if not wait_for_file_stable(path, self.logger):
            self.logger.warning(f"文件未稳定，跳过: {path}")
            return

        # 在独立线程中执行备份，不阻塞监听主线程
        t = threading.Thread(
            target=self._backup_worker,
            args=(path,),
            daemon=True,
            name=f"backup-{os.path.basename(path)[:20]}"
        )
        t.start()

    def _backup_worker(self, path: str):
        """备份工作线程"""
        try:
            self.logger.info(f"开始备份: {os.path.basename(path)}")
            ok = backup_single_file(path, self.watch_root, self.logger)
            if not ok:
                self.logger.error(f"备份失败: {path}")
        except Exception as e:
            self.logger.exception(f"备份线程异常 [{path}]: {e}")

    def on_created(self, event):
        """文件创建事件（复制粘贴、新建文件等）"""
        if not event.is_directory:
            self._handle_file(event.src_path, "CREATE")

    def on_moved(self, event):
        """
        文件移动/重命名事件。
        部分软件（如 Office、下载器）先写 .tmp 临时文件，
        完成后重命名为目标文件名，此时触发 on_moved。
        """
        if not event.is_directory:
            self._handle_file(event.dest_path, "RENAME")

    def on_modified(self, event):
        """
        文件修改事件（可选：处理覆盖写入的场景）。
        默认不处理，避免文件编辑时频繁触发备份。
        如需启用，取消下面的注释。
        """
        # if not event.is_directory:
        #     self._handle_file(event.src_path, "MODIFY")
        pass


# ─────────────────────────────────────────────────────────────────────────────
# 主入口
# ─────────────────────────────────────────────────────────────────────────────

def start_watching(logger: logging.Logger):
    """
    启动文件监听服务。

    为每个 watch_dir 创建独立的 Observer + Handler，
    支持同时监听多个目录。
    """
    observers = []

    for watch_dir in CONFIG["watch_dirs"]:
        if not os.path.exists(watch_dir):
            logger.warning(f"监听目录不存在，跳过: {watch_dir}")
            continue

        handler = MTPBackupHandler(watch_dir, logger)
        observer = Observer()
        # recursive=True 递归监听子目录
        observer.schedule(handler, watch_dir, recursive=True)
        observer.start()
        observers.append(observer)
        logger.info(f"开始监听: {watch_dir}")

    if not observers:
        logger.error("没有有效的监听目录，退出")
        sys.exit(1)

    logger.info("监听服务已启动，按 Ctrl+C 停止")
    logger.info(f"备份目标: F50\\内部共享存储空间\\{CONFIG['backup_root_name']}\\<日期>\\")

    try:
        while True:
            time.sleep(1)
            # 检查 observer 是否仍在运行
            for obs in observers:
                if not obs.is_alive():
                    logger.error("Observer 意外停止，重启...")
                    obs.start()
    except KeyboardInterrupt:
        logger.info("收到停止信号，正在关闭...")
    finally:
        for obs in observers:
            obs.stop()
        for obs in observers:
            obs.join()
        logger.info("监听服务已停止")


def main():
    parser = argparse.ArgumentParser(
        description="MTP Watch & Backup - 监听本地目录并自动备份到 F50"
    )
    parser.add_argument(
        "--daemon", action="store_true",
        help="后台运行（写 PID 文件到 watch.pid）"
    )
    args = parser.parse_args()

    logger = setup_logging(CONFIG["log_file"])

    logger.info("=" * 60)
    logger.info("MTP Watch & Backup 启动")
    logger.info(f"监听目录: {CONFIG['watch_dirs']}")
    logger.info(f"备份设备: {CONFIG['device_name']}")
    logger.info("=" * 60)

    if args.daemon:
        # 写 PID 文件
        pid = os.getpid()
        with open(CONFIG["pid_file"], "w") as f:
            f.write(str(pid))
        logger.info(f"后台模式，PID: {pid} -> {CONFIG['pid_file']}")

    start_watching(logger)


if __name__ == "__main__":
    main()
