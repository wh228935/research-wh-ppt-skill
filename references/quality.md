# 核验与交付

内容：对照原文、单位、条件、原始值、因果强度和必要限定，核查用户选择完整、图例坐标未缺失、正文与讲稿一致。不能只审措辞或文件存在。

结构：`python scripts/office.py audit final.pptx --json audit.json`检查ZIP/XML、关系、备注、可编辑对象、外链、静态效果、页外几何。结果不能证明科学结论、图内小字、字距或审美。

实际软件：`Render-Office.ps1 -Source ... -OutputDir ... -Application PowerPoint`与`-Application WPS`分别只读打开导出逐页PNG，记录来源PPT哈希、软件版本、预览哈希。分别逐页看标题/正文/图注、清晰度、遮挡溢出、间距主次和导航。未实际验证的软件明确列出；一个报告不能冒充另一个。

Word从最终PPT实际备注生成，页序/标题/讲稿对应，不能只用旧plan导出。需在Word实际检查排版；若导PDF且本机有用户指定Export-WordPdf.ps1必须用它。其他Word COM操作也遵守命名互斥锁和任务进程所有权清理。Word正文无截图、技术检查记录或强加来源列表。未检查不写已验证。

`project.py archive`输入bundle清单：最终PPT、Word、逐页PNG、缩略图、QA报告和哈希。拒绝已有v01/v02目录，记录skill版本；PPT/Word/预览必须对应同版。正式与备份数量分列。阅读报告、原图、批准、代码调试在工作目录，不混成品；默认不交PDF。

自动检查不替代逐页目视或用户反馈。报告未验证范围，不能为归档把未做的visual_reviewed填true。已确认大纲和代表页后的全稿，完成实际核验即可交付，不额外强制增加一次交付审批。
