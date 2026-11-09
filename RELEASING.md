# Windows 构建与发布指南

## 1. 版本编号

每次发布对应一个 Git 标签。`version.py` 中的版本值用于界面显示和便携包命名。创建标签时，在版本值前加 `v`。

| 版本示例 | 用途 | Releases 设置 |
| --- | --- | --- |
| `v2.1.0` | 首个公开稳定版，沿用原有 2.1 版本号 | 正式发布并设为最新版 |
| `v2.1.1` | 修复问题和兼容性 | 正式发布并设为最新版 |
| `v2.2.0-beta.1` | 新功能测试 | 勾选预发布 |
| `v2.2.0` | 新功能验收后的稳定版 | 正式发布并设为最新版 |

历史版本保留各自的标签、更新说明和下载包。同一版本的 x64 和 ARM64 包挂在同一个 Release 下。

## 2. 首次上传源码

在 GitHub 创建空仓库 `KammaTool-for-win` 并选择公开可见性。请在本公开版目录内运行以下命令。

```powershell
git init -b main
git add .
git status --short
git diff --cached --stat
git commit -m "Prepare first public version"
git remote add origin https://github.com/KAMMA415/KammaTool-for-win.git
git push -u origin main
```

提交前检查暂存清单。`config.json`、虚拟环境、缓存、打包输出和快捷方式由 Git 忽略。Git 若要求配置作者信息，请在本仓库运行下列命令并替换示例值，再重新提交。

```powershell
git config user.name "你的姓名或昵称"
git config user.email "你的 GitHub 提交邮箱"
```

首次推送按 Git 提示登录 GitHub。

## 3. 构建并验收便携包

1. 推送待发布的源码。后续版本先更新 `version.py` 和 `CHANGELOG.md`。
2. 在 GitHub 打开 **Actions → Build Windows packages → Run workflow**。默认选择 `main`，测试版可选包含该版本源码的分支。
3. 在运行结果底部的 **Artifacts** 下载 x64 和 ARM64 便携包。
4. 在对应架构的 Windows 上完整解压，确认主窗口、三个工具页、图标和基础文件操作可用。Office 导出功能在安装相应 Office 桌面应用的机器上检查。

如需在本机打包，请使用对应架构的 Python 3.13。生成的 ZIP 位于 `dist`。

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.\.venv\Scripts\python.exe build_windows.py
```

## 4. 发布到 Releases

1. 打开已通过验收的 **Build Windows packages** 运行详情，复制页面地址末尾的运行编号。
2. 进入 **Actions → Publish Windows release → Run workflow**，选择 `main`，把运行编号填入 `build_run_id` 并运行。
3. 发布流程核对两个 ZIP 的完整性及 SHA-256。它从构建时的 `version.py` 读取版本号，从 `CHANGELOG.md` 读取该版本说明，并将标签指向对应的构建提交。
4. 流程成功后打开 **Releases**，核对版本号、更新说明和两个下载包。稳定版会标为最新版，带 `alpha`、`beta` 或 `rc` 的版本会标为预发布。

GitHub 会附带源码压缩包。普通用户在 Assets 中下载对应架构的便携 ZIP。
