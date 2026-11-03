# eq2docx

[English](README.md) | 中文

**把剪贴板上的行间公式（LaTeX）用 pandoc 转成只含一条公式的 `.docx`，粘进 Word 就是原生可编辑的 OMML 公式。**

在 Word 里写论文、公式用 AXMath 敲的话，原来的循环很烦：导出 LaTeX → 手建 markdown → 跑 pandoc → 打开 docx → 选中公式 → 复制 → 切回 Word → 粘贴。`eq2docx` 把它压成一条命令。

```
AXMath  Ctrl+T  →  eq2docx -o EQ01  →  打开 EQ01.docx  →  复制公式  →  粘进 Word
```

LaTeX 是**原样**交给 pandoc 的 —— 脚本不改写、不归一化、不替你“修”公式。markdown 预览里能渲染出来的，docx 里就能渲染出来。

## 环境要求

- Windows + PowerShell。推荐 `pwsh` 7；Windows PowerShell 5.1 也能跑（脚本带 UTF-8 BOM）。
- [pandoc](https://pandoc.org/installing.html)，版本别太旧就行。已在 3.13 与 3.8 上测过，公式输出完全一致。

## 快速开始

先把文件弄到本机：把仓库 clone 下来（或下载 ZIP 解压）到你选的目录 —— 下面用 `D:\tools\eq2docx` 作示例路径。

```cmd
:: 只 clone 一次（或把 ZIP 解压到同一位置）
git clone <repo-url> D:\tools\eq2docx

:: 然后在该目录下
cd /d D:\tools\eq2docx
pwsh -File eq2docx.ps1 -Init -PandocPath "D:\tools\pandoc\pandoc.exe"
```

`-PandocPath` 可以不给：省略时 `-Init` 会从环境变量、`PATH`、常见安装位置自动探测 pandoc（见 [pandoc 是怎么找到的](#pandoc-是怎么找到的)）。

`-Init` 会做三件事：

1. 把 pandoc 路径写进 `eq2docx.config.json`，以后不用再指定；
2. 生成 `eq2docx.cmd`，这样 **cmd 里能直接敲 `eq2docx`**（cmd 不认 `.ps1`，必须靠这个壳）；
3. 把本目录加进用户 PATH，并广播 `WM_SETTINGCHANGE`。

之后**新开一个终端**才生效 —— 环境变量只对新进程有效。验证：

```cmd
eq2docx -Check
```

会打印解析到的 pandoc 路径、版本、配置文件位置。

## 日常用法

```cmd
:: 1. 把公式复制到剪贴板（AXMath: Ctrl+T，或选中整段 $$...$$ 按 Ctrl+C）
:: 2.
eq2docx -o EQ01
```

产出 `EQ01.md`（剪贴板原文，一字不改，UTF-8 无 BOM）和 `EQ01.docx`（只含那一条行间公式）。默认写在脚本目录，`-OutDir` 可改。

打开 docx，点中公式复制，粘进目标文档 —— Word 会把它变成可编辑的公式对象。

## 参数

| 参数 | 作用 |
|---|---|
| `-o <名字>` | 输出基名；`EQ01` → `EQ01.md` + `EQ01.docx`。不给则用 `EQ_yyyyMMdd_HHmmss` |
| `-OutDir <目录>` | 输出目录，默认脚本所在目录（不存在会自动建） |
| `-InputFile <文件>` | 不从剪贴板，直接读一个 md 文件（测试 / 重跑用） |
| `-PandocPath <exe>` | 本次指定 pandoc；路径无效直接报错退出，不静默回退 |
| `-Init` | 初始化：写配置 + 生成 `eq2docx.cmd` + 加 PATH |
| `-Check` | 只打印当前解析到的 pandoc 路径、版本、配置文件位置 |

## pandoc 是怎么找到的

先命中者胜：

1. `-PandocPath`（本次）
2. 环境变量 `EQ2DOCX_PANDOC`（运行时覆盖，优先于配置文件）
3. `eq2docx.config.json`
4. 环境变量 `PANDOC`
5. PATH 里的 `pandoc`
6. 常见安装位置：`%LOCALAPPDATA%\Pandoc\pandoc.exe`、`C:\Program Files\Pandoc\pandoc.exe`、`D:\anaconda3\Library\bin\pandoc.exe`

所以三种用法都成立：靠配置、`set EQ2DOCX_PANDOC=...` 临时改、或 `-PandocPath` 单次指定。

## 目录里有什么

| 文件 | 说明 |
|---|---|
| `eq2docx.ps1` | 主脚本 |
| `README.md` | 英文说明（`README.zh-CN.md` 为中文版） |
| `eq2docx.cmd` | `-Init` 生成的 cmd 壳，别手改 |
| `eq2docx.config.json` | `-Init` 写入的 pandoc 路径 |
| `demo_equation.md` | 示例输入：cases + 嵌套矩阵，AXMath 原样导出 |
| `demo_equation.afx` | 对应的 AXMath 源文件 |

## 退出码

| 码 | 含义 |
|---|---|
| 0 | 成功 |
| 1 | `-Init` 没探测到 pandoc |
| 2 | 找不到 pandoc / `-PandocPath` 无效 |
| 3 | 输入为空，或输出名含非法字符，或输入文件不存在 |
| 4 | pandoc 转换失败 |

## 脚本自带的自检

转完后脚本会解压 docx，检查 `word/document.xml` 里有没有 `<m:oMath`：

- `√ 已生成 Office 公式（OMML）` —— 公式是真数学对象。
- `! docx 里没检测到 OMML 公式` —— 内容没被 `$...$` 包住，pandoc 当普通文本处理了。docx 照样生成，只是里面是一段 LaTeX 源码。

## 常见问题

**cmd 里 `eq2docx` 提示不是内部或外部命令**
没跑 `-Init`，或者跑完没重开窗口（PATH 广播只对新进程生效）。

**控制台输出乱码**
控制台代码页和 PowerShell 的编码不一致。`.md` 文件本身是 UTF-8，不受影响；`chcp 936`（或 `chcp 65001`）让两者对齐即可。

**docx 里只有公式、没有正文**
设计如此：输入是一条行间公式，输出就是那一条。要整篇文档就直接调 pandoc —— `pandoc your.md -o your.docx`。

**公式在 Word 里显示成灰色 LaTeX 文本**
交给 pandoc 的内容没被 `$$...$$` 包住，或者 `$$` 被换行截断了。脚本会告警，但不会替你补。

**换电脑**
拷 `eq2docx.ps1`（需要示例就连 `demo_equation.md` 一起拷），在新机器上跑一次 `-Init -PandocPath "<本机 pandoc.exe>"`，`.cmd` 和配置会自动重建。从网络共享或 U 盘拷来的文件先 `Unblock-File`，否则 RemoteSigned 会拦。

## 许可证

MIT，见 [LICENSE](LICENSE)。
