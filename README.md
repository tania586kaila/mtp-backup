# MTP Backup Tool

> 将 Windows 本地文件夹通过 **MTP 协议**自动备份到中兴 F50（或其他 MTP 设备）内部存储，按日期命名备份目录，支持增量备份。

---

## 功能特性

| 功能 | 说明 |
|------|------|
| MTP 直传 | 使用 Windows Shell COM 接口，无需第三方驱动 |
| 中文文件名 | 通过 `ParseName` 获取 FolderItem，完整支持中文路径 |
| 按日期归档 | 每次备份自动创建 `YYYYMMDD` 格式子目录 |
| 增量备份 | 已存在的文件自动跳过，不重复传输 |
| 多源目录 | 支持同时备份多个本地文件夹 |
| 目录结构保留 | 递归复制，保持原有子目录层级 |
| 完整日志 | 控制台 + 文件双输出，UTF-8 编码，记录每个文件操作 |

---

## 备份目录结构（F50 侧）

```
F50\内部共享存储空间\
└── PC备份\                    ← backup_root_name（可配置）
    └── 20260316\              ← 按日期自动命名（YYYYMMDD）
        └── sync\              ← 与本地源目录同名
            ├── file1.docx
            ├── file2.xlsx
            └── subdir\
                └── file3.png
```

---

## 环境要求

- **操作系统**：Windows 10 / 11
- **Python**：3.8+
- **设备**：中兴 F50（或任意支持 MTP 的移动设备），通过 USB 连接并在"此电脑"中可见
- **依赖包**：`pywin32`

---

## 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/your-username/mtp-backup.git
cd mtp-backup
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 修改配置

打开 `mtp_backup.py`，修改顶部 `CONFIG` 字典：

```python
CONFIG = {
    # 本地源文件夹路径（支持多个）
    "source_dirs": [
        r"F:\share\sync",          # 修改为你的路径
        r"C:\Users\你的用户名\Documents",
    ],

    # F50 设备名称（模糊匹配）
    "device_name": "F50",

    # F50 上的备份根目录名
    "backup_root_name": "PC备份",

    # 日期格式（strftime 格式）
    "date_format": "%Y%m%d",

    # 单文件传输超时（秒）
    "copy_timeout_sec": 60,

    # 日志文件路径
    "log_file": r"C:\path\to\backup.log",
}
```

### 4. 连接设备并运行

确保 F50 已通过 USB 连接，在"此电脑"中可见，然后：

```bash
python mtp_backup.py
```

---

## 运行示例

```
[2026-03-16 21:59:50] INFO  - ============================================================
[2026-03-16 21:59:50] INFO  - MTP 备份开始
[2026-03-16 21:59:50] INFO  - ============================================================
[2026-03-16 21:59:50] INFO  - 找到设备: F50
[2026-03-16 21:59:50] INFO  - 存储空间: 内部共享存储空间
[2026-03-16 21:59:50] INFO  - 备份目标: F50\内部共享存储空间\PC备份\20260316
[2026-03-16 21:59:50] INFO  - 发现 6 个文件
[2026-03-16 21:59:52] INFO  - [COPY] 文件1.docx
[2026-03-16 21:59:54] INFO  - [OK]   文件1.docx
[2026-03-16 21:59:54] INFO  - [COPY] 文件2.xlsx
[2026-03-16 21:59:56] INFO  - [OK]   文件2.xlsx
...
[2026-03-16 22:00:10] INFO  - ============================================================
[2026-03-16 22:00:10] INFO  - 备份完成！[OK] 6  [SKIP] 0  [FAIL] 0
[2026-03-16 22:00:10] INFO  - ============================================================
```

---

## 项目结构

```
mtp-backup/
├── mtp_backup.py       # 主程序（含完整注释）
├── mtp_backup.ps1      # PowerShell 版本（备用，直接内联执行）
├── requirements.txt    # Python 依赖
├── .vscode/
│   └── launch.json     # VS Code 调试配置（F5 一键运行）
├── .gitignore
└── README.md
```

---

## 技术原理

### MTP 访问方式

Windows 通过 **WPD（Windows Portable Device）** 框架管理 MTP 设备。本工具使用 **Shell.Application COM 接口**（`Shell32.dll`）操作 MTP 设备，这是最兼容的方式：

```
Shell.Application
  └── Namespace(0x11)          # 此电脑
      └── F50 设备
          └── 内部共享存储空间
              └── PC备份/
                  └── 20260316/
                      └── sync/
```

### 关键 API

| API | 用途 |
|-----|------|
| `Shell.Namespace(0x11)` | 获取"此电脑"命名空间 |
| `Folder.Items()` | 枚举目录内容 |
| `Folder.NewFolder(name)` | 创建子目录 |
| `Namespace(path).ParseName(file)` | 获取文件的 FolderItem（支持中文） |
| `Folder.CopyHere(item, 0x14)` | 静默复制文件到 MTP 目录 |

### 中文文件名处理

直接用路径字符串操作 MTP 设备时，中文文件名容易出现乱码或找不到文件的问题。本工具通过 `Shell.Namespace(dir).ParseName(filename)` 获取 `FolderItem` 对象，再传给 `CopyHere`，绕开了路径字符串编码问题。

### 增量备份

每次传输前调用 `Folder.Items()` 获取目标目录现有文件名集合，若文件名已存在则跳过，实现增量备份。

---

## 常见问题

**Q: 提示"未找到设备 F50"**
A: 确认 F50 已通过 USB 连接，并在 Windows"此电脑"中以 MTP 模式显示（不是 U 盘模式）。

**Q: 文件传输超时**
A: 增大 `copy_timeout_sec` 配置值；或检查 USB 连接是否稳定。

**Q: 日志文件乱码**
A: 日志文件使用 UTF-8 编码，用支持 UTF-8 的编辑器（如 VS Code）打开。

**Q: 能否定时自动备份？**
A: 可以用 Windows 任务计划程序定时执行 `python mtp_backup.py`。

---

## License

MIT
