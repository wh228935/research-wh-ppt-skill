# 安装到 Codex

仓库名 `research-wh-ppt-skill`，skill 名与安装目录名 `research-wh-ppt-skil`。这两个名称按创建者的选择分别保留。

## 当前状态

当前是 `draft/initial-skill` 分支的开发版，尚未合并到 main、尚未发布稳定版本。创建者正在审阅代表页，尚未正式安装本机个人副本。不要把开发版称为正式成品。

创建者已决定公开仓库；公开设置完成后无需 GitHub 协作者权限即可下载。若访问仍要求登录或提示404，先核实实际可见性与网址。

## 稳定版发布后的推荐方法

将仓库网址和这段文字发给 Codex：

> 请使用 skill-installer 安装 https://github.com/wh228935/research-wh-ppt-skill 。skill 在 main 分支的仓库根目录，路径为 .，安装名为 research-wh-ppt-skil。先检查本机是否已有同名 skill，已有时报告版本并提出更新方案，不能直接覆盖。

使用官方 skill-installer 脚本时，参数为：

```text
python <skill-installer目录>/scripts/install-skill-from-github.py --repo wh228935/research-wh-ppt-skill --ref main --path . --name research-wh-ppt-skil
```

`<skill-installer目录>` 由当前 Codex 环境定位，不是要原样输入的路径。仓库根目录的 `--path .` 需配合显式 `--name`，避免从“.”推导安装名。安装目标使用当前宿主个人 skills 目录；一般为 `$CODEX_HOME/skills`，未设置时为 `~/.codex/skills`。

安装完成后，在下一轮对话调用 `$research-wh-ppt-skil`，例如“使用这个 skill，根据我提供的论文准备科研汇报”。依赖检查通过后才开始相应制作任务。

## 开发版测试

普通使用者等待稳定版。开发者明确选择测试开发分支时，用相同官方安装器，将 `--ref` 改为 `draft/initial-skill`，并将 `--dest` 指向单独测试目录；不要覆盖个人稳定副本。分支名含斜杠，优先使用 `--repo` 和 `--ref`，避免把 tree URL 的路径解析错。

开发版测试不等于稳定发布，官方通用安装器也不代替本项目的发布确认。项目自带的 `scripts/Install-Skill.ps1` 用于确认后的稳定包首次安装，开发状态会拒绝执行。

## 必要环境

- Windows、PowerShell 7、Python 3（核心脚本使用标准库）。
- PowerPoint 用于原生 PPTX 制作与主要渲染检查；本机同时有 WPS 时补充兼容性检查。
- Word 用于讲稿的原生显示检查或PDF校验；PDF只是内部检查文件，不是默认成品。
- 读取PDF时需要可用的PDF提取和渲染工具。若环境缺失，先说明缺项、用途和安装方案，取得使用者同意后再安装。

不依赖创建者的其他私人 skills，不捆绑 Office、字体安装包、原始研究PPT或论文图件。仅支持Windows；安装工作流本身不会让未安装的应用自动可用。

## 更新和回退

安装器遇到已有同名目录会停止。更新前先检查现有版本、说明改动并保留旧包，按使用者确认的方案操作。制作中的PPT默认继续沿用原skill版本与已确认风格。创建者本机副本不随GitHub提交自动更新。
