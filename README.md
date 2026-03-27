# Excel 图片批量下载（姓名+手机号分文件夹）

## 功能

- 读取 `摩托车消费券核销记录.xlsx`
- 对每一行：用 **B 列（姓名）+ C 列（手机号）** 作为文件夹名
- 下载 **L 到 T 列** 单元格内的所有图片链接
  - 同一单元格多张图片用 `|` 分隔
- 图片保存到输出目录下的独立文件夹中（每人一个文件夹）

## 本地运行（Python 3）

建议使用虚拟环境：

```bash
python3 -m venv .venv
./.venv/bin/python -m pip install -U pip
./.venv/bin/python -m pip install -r requirements.txt
./.venv/bin/python download_excel_images.py
```

默认会读取当前目录的 `摩托车消费券核销记录.xlsx`，输出到 `下载图片/`。

常用参数：

```bash
./.venv/bin/python download_excel_images.py --out "下载图片" --start-row 3 --skip-existing
```

如你的网络必须走系统代理（公司代理等），可加上 `--use-proxy`：

```bash
./.venv/bin/python download_excel_images.py --use-proxy
```

## 打包（Windows / macOS）

注意：**PyInstaller 需要在目标系统分别打包**（Windows 上打 Windows 可执行文件，macOS 上打 macOS 可执行文件）。

### 1) 安装 PyInstaller

在对应系统的虚拟环境中执行：

```bash
python -m pip install -U pyinstaller
```

### 2) 打包命令

在项目目录执行：

```bash
pyinstaller --onefile --name download_excel_images download_excel_images.py
```

打包产物在 `dist/`：

- Windows：`dist\\download_excel_images.exe`
- macOS：`dist/download_excel_images`

### 3) 运行打包产物

把可执行文件和 `摩托车消费券核销记录.xlsx` 放到同一目录，执行：

```bash
./download_excel_images --xlsx "摩托车消费券核销记录.xlsx" --out "下载图片"
```

## 说明

- 默认从第 3 行开始处理（第 1 行标题、第 2 行表头）。如你的表格不同，可用 `--start-row` 调整。
- 文件命名格式：`列名_序号.扩展名`（列名优先取第 2 行表头，缺失时用列字母 L~T）。

