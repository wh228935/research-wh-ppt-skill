# 科研PPT Skill

Windows独立工作流：论文精读与逐章确认、原图选择、详细大纲、3—5张代表页、可编辑PPT、详细备注和Word讲稿。支持局改和全稿美化。

调用`$research-wh-ppt-skil`；默认可用于科研PPT，用户指定其他skill时尊重其选择。

GitHub仓库名为`research-wh-ppt-skill`，skill文件夹和调用名为`research-wh-ppt-skil`，依用户分别指定的名称保留。

## 用 Codex 安装

把这个网址发给 Codex：

**https://github.com/wh228935/research-wh-ppt-skill**

直接复制这句话即可：

```text
请使用 skill-installer 安装这个科研PPT skill：
https://github.com/wh228935/research-wh-ppt-skill
skill 位于 main 分支的仓库根目录，路径为 .，安装名为 research-wh-ppt-skil。
安装到我的个人 skills 目录；已有同名目录时先告知我，不直接覆盖。
```

安装后，在下一轮对话中输入 `$research-wh-ppt-skil`，再提供论文Word/PDF和汇报要求。

### 给 Codex 的安装定位信息

- 仓库：`wh228935/research-wh-ppt-skill`
- 分支：`main`
- skill路径：`.`（仓库根目录有 `SKILL.md`）
- 安装名：`research-wh-ppt-skil`（保留作者指定的拼写）
- 使用宿主自带的 `skill-installer`，参数为 `--repo wh228935/research-wh-ppt-skill --ref main --path . --name research-wh-ppt-skil`。
- 如果宿主有不同的个人skills目录，以该宿主的安装工具为准，不照搬创建者电脑路径。

完整步骤见 [INSTALL.md](INSTALL.md)。公开仓库可匿名下载，不需要创建者账号或令牌。

## 当前版本

这是可安装的**公开预览版**，版本记录为 `unreleased`；代表页已经试做，最终排版审阅和稳定版本号尚未确定。公开预览安装不等于稳定版确认，不声称全部论文类型或Office版本均已验证。

已完成的技术检查与边界见 [技术验证记录](validation/TECHNICAL_REPORT.md)。本仓库仅包含工作流、主题参数和制作工具，未包含创建者的完整论文、研究PPT或论文原图。

## 环境和运行

PowerShell 7中先运行`python scripts/preflight.py`。核心Python标准库；PPT原生制作/渲染需本机Office；可选PDF工具缺失先报告，经同意后安装。读SKILL.md开始。工作文件放任务目录，不写本包。

复制完整文件夹到目标宿主的个人skills目录，先检查是否已有未知版本，不能覆盖。本机会话指定目录优先，不写死通用绝对路径。首次安装用已验证并确认的包，后续不自动更新本机副本。

工具：project.py状态/确认/归档；office.py材料提取/结构检查/讲稿；Build-Deck.ps1原生制作；Render-Office.ps1只读预览。各工具--help或Get-Help查看入口。

版本管理采用独立分支、验证、用户确认后合并/标版本。PPT稿次与skill版本分开。不会打包整份研究PPT、研究结论或学校身份。
