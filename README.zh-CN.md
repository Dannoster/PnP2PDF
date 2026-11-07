# PnP2PDF

[![tests](https://github.com/Dannoster/PnP2PDF/actions/workflows/tests.yml/badge.svg)](https://github.com/Dannoster/PnP2PDF/actions/workflows/tests.yml)

[English](README.md) · [Русский](README.ru.md) · **简体中文** · [Español](README.es.md) · [Deutsch](README.de.md)

PnP2PDF 把 print-and-play 卡牌图片排版成一个可以直接打印的 PDF。选择放卡图的
文件夹、纸张和网格，就能得到一个带裁切线的 PDF；需要时还会把卡背镜像排到背面。

- 只有一个窗口，不需要命令行；
- 打印前可以先看实时预览；
- 支持无卡背、多个卡牌共用一个卡背、多个不同卡背，以及每张卡自带卡背的情况；
- 默认值适配 A4 纸张和 62.5 × 88.7 毫米（扑克牌尺寸）的卡牌。

## 目录

- [免安装版本](#免安装版本)
- [从源码运行](#从源码运行)
- [快速开始](#快速开始)
- [卡牌文件夹结构](#卡牌文件夹结构)
- [设置](#设置)
- [示例](#示例)
- [命令行](#命令行)
- [打包独立应用](#打包独立应用)
- [出问题时](#出问题时)
- [开发](#开发)
- [已知问题](#已知问题)
- [许可协议](#许可协议)

## 免安装版本

在 [Releases](https://github.com/Dannoster/PnP2PDF/releases) 下载对应系统的压缩包，
解压后直接运行，无需安装 Python。

macOS 版本没有使用 Apple Developer ID 签名，所以第一次运行下载的副本时需要去掉
隔离属性：

```bash
xattr -dr com.apple.quarantine /Applications/PnP2PDF.app
open /Applications/PnP2PDF.app
```

## 从源码运行

需要 Python 3.11 或更新版本。

```bash
git clone https://github.com/Dannoster/PnP2PDF.git
cd PnP2PDF
pip install -r requirements.txt
python PnP2PDF.py
```

安装为软件包后会得到两个命令：`pnp2pdf` 打开窗口，`pnp2pdf-cli` 不开窗口直接生成。

```bash
pip install .
pnp2pdf
```

## 快速开始

1. **Cards folder** — 选择放卡图的文件夹。选好后右侧的预览会立即出现。
2. **纸张与卡牌尺寸** — 默认 A4（210 × 297 毫米）和 62.5 × 88.7 毫米的卡牌。勾选
   *Smaller size (98%)* 会把卡牌印得略小，卡与卡之间留出更多纸。
3. **排版** — 行数 × 列数、卡牌之间的间距，以及可选的不可打印边距（安全区）。
4. **Double sided cards** — 文件夹里放了卡背时勾选它。详见
   [卡牌文件夹结构](#卡牌文件夹结构)：不勾选时，卡背图片会被当作普通卡面打印。
5. **Create PDF** — 进度条走完后窗口不会关闭，状态栏变绿并显示文件名，
   **Show PDF** 会在访达、资源管理器或文件管理器中定位该文件。

PDF 会写入卡牌目录下以时间命名的新文件夹。勾选 *Save the PDF to a different
folder* 可以改到其他位置保存。

## 卡牌文件夹结构

*卡牌文件夹* 就是你在窗口里选的那个文件夹，下面的内容都放在它里面。支持
`.png`、`.jpg`、`.jpeg`、`.tif`、`.tiff`。

图片会被拉伸填满卡牌区域，所以请使用与卡牌相同长宽比的图片。

### 没有卡背的卡牌

把图片放进文件夹即可，不需要勾选任何选项：

```
my-cards/
├── card-01.png
├── card-02.png
└── card-03.png
```

### 所有卡牌共用一个卡背

卡背是一张文件名以毫米尺寸结尾的图片，对应的卡牌放在与卡背同名（去掉尺寸部分）
的子文件夹里：

```
my-cards/
├── red-back_63,5_88,9.png     ← 卡背，名字里带尺寸
└── red-back/                  ← 使用这个卡背的卡牌
    ├── card-01.png
    └── card-02.png
```

勾选 **Double sided cards** 再生成。每一张正面之后都会跟一张同样的、镜像过的
背面，双面打印时正好对齐。

### 多个卡背

按同样的规则重复：一个卡背文件配一个同名文件夹。不同卡背的卡牌尺寸可以不同，
会排在同一套版面里。

```
my-cards/
├── red-back_63,5_88,9.png
├── blue-back_64,5_91,3.png
├── red-back/
│   └── card-01.png …
└── blue-back/
    └── card-01.png …
```

尺寸后缀可以省略：省略时卡背会被拉伸到卡牌尺寸；给出时，比卡牌大的卡背会按卡牌
窗口裁切，出血和套准标记就不会印到卡牌上。小数部分用逗号或点都可以：
`63,5` 与 `63.5` 等价。

### 每张卡自带卡背

如果卡背只属于某一张卡，把两张图都放进名为 `Double Sided` 的文件夹。程序按字母序
把文件两两配对：正面、背面、正面、背面……

```
my-cards/
└── Double Sided/
    ├── 1-front.png
    ├── 2-back.png
    └── 3-front.png …
```

文件数为奇数时会直接报错并说明原因，避免配对错位。

### 混合使用时的单面卡牌

把它们放进 `Single Sided` 文件夹：

```
my-cards/
├── Double Sided/…
└── Single Sided/
    ├── card-01.png
    └── card-02.png
```

四种结构可以放在同一个文件夹里，一次生成。

> **别忘了勾选。** 文件夹结构说明的是*哪个卡背属于哪张卡*，而
> **Double sided cards** 复选框才决定要不要打印这些卡背。不勾选时，卡背图片会被
> 当成普通卡面，和卡牌一起排到页面上。

## 设置

| 设置 | 默认值 | 说明 |
| --- | --- | --- |
| 纸张尺寸 | 210 × 297 毫米 | A4，可用任意毫米尺寸。 |
| 卡牌尺寸 | 62.5 × 88.7 毫米 | 扑克牌尺寸，图片会拉伸到这个矩形。 |
| Smaller size (98%) | 关闭 | 每张卡按上述尺寸的 98% 打印。 |
| 行 × 列 | 3 × 3 | 每张纸上的卡牌数量。 |
| 间距 | 2 毫米 | 卡牌之间的距离，裁切线画在这里。 |
| 上下安全区 | 0 毫米 | 打印机无法覆盖的边距。 |
| 左右安全区 | 0 毫米 | 水平方向的同一设置。 |
| 卡牌文件夹 | — | 上文描述的文件夹。 |
| PDF 保存位置 | 卡牌旁边 | 可选项，可以另选文件夹。 |
| Double sided cards | 关闭 | 打印卡牌背面。 |

设置会保存在应用数据目录下的 `presets.json` 中（路径见
[出问题时](#出问题时)）。从源码运行时该文件位于项目目录下。

## 示例

[`examples/`](examples/) 里有每种结构的现成卡牌集（每个只有几 KB），以及用于重新
生成的 `make_examples.py`。更多说明见 [`examples/README.md`](examples/README.md)。

| 文件夹 | 演示内容 | 页数 |
| --- | --- | --- |
| `single-sided/` | 九张无卡背卡牌 | 1 |
| `one-back/` | 六张卡共用一张卡背 | 2 |
| `two-backs/` | 两个卡背，其中一个比卡牌大 | 4 |
| `double-sided/` | 成对卡牌与单面卡牌混合 | 2 |

```bash
python3 -m pnp2pdf.cli --cards examples/one-back --out /tmp/ex --double-sided --quiet
```

## 命令行

`pnp2pdf-cli` 不开窗口即可生成同样的 PDF，适合脚本和批量处理。

```bash
pnp2pdf-cli --cards ~/cards --out ~/pdf
pnp2pdf-cli --cards ~/cards --out ~/pdf --double-sided
pnp2pdf-cli --cards ~/cards --preview /tmp/first-page.png
```

| 参数 | 含义 |
| --- | --- |
| `--cards DIR` | 卡图文件夹（必填）。 |
| `--out DIR` | PDF 输出目录，默认放在卡牌旁边。 |
| `--paper WxH` | 纸张尺寸（毫米），默认 `210x297`。 |
| `--card WxH` | 卡牌尺寸（毫米），默认 `62.5x88.7`。 |
| `--grid RxC` | 行 × 列，默认 `3x3`。 |
| `--gap MM` | 卡牌间距（毫米），默认 `2`。 |
| `--safe-ud MM` | 上下安全区（毫米），默认 `0`。 |
| `--safe-lr MM` | 左右安全区（毫米），默认 `0`。 |
| `--smaller-size` | 按给定尺寸的 98% 打印。 |
| `--double-sided` | 打印卡背。 |
| `--preview PATH` | 只把第一页写成 PNG 后退出。 |
| `--quiet` | 不输出进度。 |
| `--version` | 显示版本后退出。 |

退出码：`0` 成功，`1` 意外错误，`2` 输入有误，`130` 已取消。

## 打包独立应用

一条命令即可打包，并自动修复环境：

```bash
./build_app.sh          # macOS 与 Linux
build_app.bat           # Windows
```

脚本会创建缺失的 `.venv`、安装或更新依赖、替换损坏或来自其他平台的虚拟环境、
运行测试并打包应用。产物：

- macOS：`dist/PnP2PDF.app`
- Windows：`dist/PnP2PDF.exe`
- Linux：`dist/PnP2PDF`

`--no-tests` 跳过测试，`PNP2PDF_BUILD_DRY=1` 只检查环境、不打包。macOS 上会在
写入版本号之后重新签名，签名校验不通过时构建会给出明确错误并失败。

## 出问题时

**打包后的应用一闪就退出或没有反应。** 窗口版程序平时会隐藏的错误现在会写进
`crash.log`，错误对话框里会显示具体路径。反馈问题时请附上该文件。

- macOS：`~/Library/Application Support/PnP2PDF/crash.log`
- Windows：`%APPDATA%\PnP2PDF\crash.log`
- Linux：`$XDG_DATA_HOME/pnp2pdf/crash.log`（或 `~/.local/share/pnp2pdf/crash.log`）

**提示 “PnP2PDF is damaged and can't be opened”（macOS）。** 下载的未签名应用被
隔离了 —— 见 [免安装版本](#免安装版本)。

**提示 “No card images found”。** 图片扩展名必须是 `.png`、`.jpg`、`.jpeg`、
`.tif` 或 `.tiff`；带卡背的卡牌要放在以该卡背命名的子文件夹里。

**卡背被当成卡牌排到了页面上。** 没有勾选 **Double sided cards** —— 见
[卡牌文件夹结构](#卡牌文件夹结构)。

**应用显示成一个文件夹（macOS）。** 应用包或启动数据库过期：重新执行
`./build_app.sh`，然后运行
`/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -f dist/PnP2PDF.app`。

## 开发

```bash
pip install -r requirements-dev.txt
pytest
```

测试用替身代替 CustomTkinter，因此不需要显示器。包结构：

```
pnp2pdf/
├── app.py          事件循环与启动
├── cards.py        扫描文件夹、把卡牌与卡背配对
├── layout.py       把卡牌排到纸张上
├── render.py       绘制页面：卡牌、裁切线、角标、安全区
├── pdf.py          写出 CMYK PDF
├── preview.py      预览面板用的低分辨率快速渲染
├── job.py          一次构建：卡牌 → 版面 → PDF
├── cli.py          命令行入口
└── ui/             窗口、表单辅助、工作线程
```

## 已知问题

- Windows 版本偶尔在关闭窗口后报错，此时 PDF 已经写好。
- Microsoft Defender 有时会把 PyInstaller 打包的程序标记为未知软件。
- macOS 版本没有签名，因此需要上面的隔离处理；签名与公证需要 Apple
  Developer ID。

## 许可协议

[MIT](LICENSE)。
