# 京东商品上下架与风控状态自动化巡检系统

基于 **Python + DrissionPage + OpenPyXL** 构建的电商数据自动化巡检 RPA 工具。实现从 Excel 表格自动读取京东商品链接、原生指纹仿真移动端 H5 绕过风控、多分支识别商品在售/下架状态，并实时回写结果至 Excel，支持断点续传、直连防风控与一键绿色打包。

---

## 🌟 跨平台支持与核心功能

* **跨平台兼容**：全面兼容 **Windows 10/11** 与 **macOS (Mac)**，自动识别本地 Chrome 与 Edge 浏览器路径。
* **原生防风控（0验证码、0死滑块）**：
  * 采用浏览器 100% 真实指纹（移除假 UA 避免 Sec-CH-UA 矛盾）；
  * 默认开启 `--no-proxy-server` 网络直连，强制绕过本机 VPN/Clash 代理，彻底杜绝境外机房 IP 触发死滑块。
* **智能断点续传**：启动时自动扫描已有结果，跳过已标记为有效状态（`0` 或 `1`）的行；支持只针对“待查/异常”行进行重新补测。
* **原表格保护**：使用 `openpyxl(data_only=False)`，保留原表格的所有计算公式、背景颜色与图片排版。

---

## 📦 如何打包发给其他电脑使用？

### 场景 1：发给 Windows 电脑同事（免装 Python 绿色版）

如果对方电脑没有安装 Python 环境，你可以直接打包为一个**免安装绿色便携包**发给他们：

1. **一键生成打包文件**：
   * 在本项目根目录下，双击运行 **[`package.bat`](file:///C:/Users/q1894/Desktop/code/PYcode/jd_status_checker/package.bat)**；
   * 程序会自动调用 PyInstaller 进行完整封装，耗时约 30 秒。
2. **打包成果物**：
   * 生成的免安装绿色程序存放在 **`dist\京东商品状态巡检工具`** 文件夹中；
   * 该文件夹内包含了运行所需的所有环境，并且自带 `京东商品状态巡检工具.exe`。
3. **分发方式**：
   * 你只需要将 `dist\京东商品状态巡检工具` 整个文件夹**压缩为 `.zip` 压缩包**；
   * 发送给同事（通过微信、企业微信或网盘）；
   * **对方收到后只需右键解压，双击里面的 `京东商品状态巡检工具.exe` 即可直接使用！对方电脑无需安装任何 Python 环境！**

---

### 场景 2：发给 Mac (macOS) 电脑同事

由于 Windows 生成的 `.exe` 无法直接在 Mac 上运行（不同操作系统的底层二进制不同），发给 Mac 用户有两种极为简便的方式：

#### 方案 A：直接发项目源码（最推荐，Mac 用户使用最快）
Mac 几乎全部自带 Python 3 环境，并且使用终端非常方便：
1. 将当前项目打包为 `.zip`（或让同事直接通过 Git 下载：`git clone https://github.com/BoBo-v/jd_status_checker.git`）；
2. Mac 同事解压后，打开终端进入该目录，运行：
   ```bash
   chmod +x run.sh
   ./run.sh
   ```
3. `run.sh` 脚本会自动检测并补齐所需依赖，并弹出交互向导，自动调用 Mac 上的 Chrome 或 Edge 运行巡检！

#### 方案 B：在 Mac 上一键打包为 Mac 专属可执行文件
如果有同事手头有 Mac 电脑，希望打出一个让所有 Mac 小白双击即用的免 Python 包：
1. 在那台 Mac 的项目目录下运行：
   ```bash
   chmod +x package.sh
   ./package.sh
   ```
2. 打包完成后将在 `dist/京东商品状态巡检工具_Mac` 下生成 Mac 原生免环境程序，压缩后发给其他 Mac 用户即可双击使用！

---

## 🚀 日常运行说明

### Windows 系统：
* 双击运行 **`run.bat`** 即可。

### Mac / Linux 系统：
* 在终端执行 **`./run.sh`** 即可。

### 命令行常用参数：
```bash
# 指定 Excel 路径运行
python main.py --excel "你的商品表格.xlsx"

# 无头后台静默运行
python main.py --excel "你的商品表格.xlsx" --headless

# 仅重测待查与异常数据
python main.py --excel "你的商品表格.xlsx" --recheck-errors

# 强制全部重新检测
python main.py --excel "你的商品表格.xlsx" --rerun-all
```
