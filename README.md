# 科研PPT Skill

Windows独立工作流：论文精读与逐章确认、原图选择、详细大纲、3—5张代表页、可编辑PPT、详细备注和Word讲稿。支持局改和全稿美化。

调用`$research-wh-ppt-skil`；默认可用于科研PPT，用户指定其他skill时尊重其选择。

GitHub仓库名为`research-wh-ppt-skill`，skill文件夹和调用名为`research-wh-ppt-skil`，依用户分别指定的名称保留。

## 开发状态

未发布稳定版，尚待用户确认代表页与版本号。VERSION.json用unreleased；检查通过不替代原文核对和逐页目视。

已完成检查与边界见[技术验证记录](validation/TECHNICAL_REPORT.md)。当前不安装开发分支来接管日常任务。

## 环境和运行

PowerShell 7中先运行`python scripts/preflight.py`。核心Python标准库；PPT原生制作/渲染需本机Office；可选PDF工具缺失先报告，经同意后安装。读SKILL.md开始。工作文件放任务目录，不写本包。

复制完整文件夹到目标宿主的个人skills目录，先检查是否已有未知版本，不能覆盖。本机会话指定目录优先，不写死通用绝对路径。首次安装用已验证并确认的包，后续不自动更新本机副本。

工具：project.py状态/确认/归档；office.py材料提取/结构检查/讲稿；Build-Deck.ps1原生制作；Render-Office.ps1只读预览。各工具--help或Get-Help查看入口。

版本管理用私有仓库、独立分支、验证、用户确认后合并/标版本。PPT稿次与skill版本分开。不会打包整份研究PPT、研究结论或学校身份。
