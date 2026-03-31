# Excel 图片批量下载（PySide6 GUI）

## 功能

- 图形界面运行，支持选择 Excel 文件和输出目录
- 每行使用 **行号** 生成文件夹：`1`、`2`、`3`...（更通用，不依赖姓名/手机号列）
- 不再固定列范围，改为遍历每行**所有单元格**
- 单元格内支持多链接分隔：`|`、`,`、`，`、`;`、换行
- 仅下载 `http://` 或 `https://` 链接
- 下载时显示进度（已完成/总数/百分比）
- 支持“停止下载”
- 支持“继续下载（跳过已下载）”
- 结束后显示：总数、成功数、失败数、失败行号

## 运行环境

- Python 3.9+
- 依赖：`pyside6`、`requests`、`openpyxl`

安装依赖：

```bash
python3 -m pip install -U pip
python3 -m pip install pyside6 requests openpyxl
```

## 本地运行

在项目目录执行：

```bash
python3 download_excel_images.py
```

默认会启动 GUI。

## 命令行模式（可选）

如需命令行运行：

```bash
python3 download_excel_images.py --xlsx "摩托车消费券核销记录.xlsx" --out "下载图片"
```

常用参数：

- `--start-row 3`：从第几行开始处理
- `--skip-existing`：跳过已存在文件
- `--use-proxy`：使用系统代理
- `--gui`：强制启动图形界面

## 打包

注意：PyInstaller 不能跨平台直接出可执行文件，需在目标系统打包。

### macOS 本地打包（mac 可执行）

```bash
pyinstaller --onefile --windowed --name DownloadPicFromExcel --collect-all PySide6 download_excel_images.py
```

产物：`dist/DownloadPicFromExcel`

### Windows 打包（出 .exe）

在 Windows 环境执行：

```bash
pyinstaller --onefile --windowed --name DownloadPicFromExcel --collect-all PySide6 download_excel_images.py
```

产物：`dist\DownloadPicFromExcel.exe`

## GitHub Actions 自动打 Windows 包

仓库已包含工作流：

- `.github/workflows/build-windows.yml`

推送到 GitHub 后，在 Actions 执行 `Build Windows EXE`，可下载 artifact：

- `DownloadPicFromExcel-Windows.zip`

## 说明

- 默认从第 3 行开始处理（通常第 1 行标题、第 2 行表头）
- 下载文件命名格式：`列名_序号.扩展名`（列名优先取第 2 行表头，缺失时用列字母）

