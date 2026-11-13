# 阅读、规划与确认

任务目录保存project.json、intake.json、阅读报告、证据清单、图件缩略图及选择、内容取舍、plan.json、prototype-selection.json、审核文件。原图和源文档留工作区，不放进skill包。

## 必须由真实用户回复确认的对象

| scope | 对应文件 | 内容 |
|---|---|---|
| intake | intake.json | 时长、背景、材料、场景选择，阅读范围 |
| reading-overview:文档ID | 整体理解报告 | 研究问题、材料、方法、结果、结论与限定 |
| reading-chapter:文档ID:章ID | 每章报告 | 逐章完整理解，含图表及原文定位，分别确认 |
| figures | 图件清单与选择 | 缩略图、作用、用户选择、拟用页 |
| content | 内容取舍报告 | 必讲建议、删减选择、贡献局限与分析区分 |
| outline | 完整plan.json | 每页上屏文字、图件、布局、重点、预计用时 |
| prototype-selection | pages列表JSON | 3—5张不同内容代表页，先确认清单再试做 |
| prototypes | 审核文件 | 代表页PPT、预览和风格参数的文件指纹；最终方案确认 |
| full-review（若用户审阅全稿） | 审核文件 | 最终PPT、Word、预览、QA的指纹和用户回复；不额外强制正式交付前再等一次批准 |

`project.py record --scope ... --file ... --quote ...`记录真实回复；文件发生变化后旧确认不再有效。可以继续不依赖该确认的阅读与整理，不得越过待确认环节去完成全稿。不能创建模板报告后批量填“已精读/已确认”。

intake.documents每篇包含id、path和chapters列表（每章id/title）。正文及正文图表全部精读；附录补充先说明关联让用户选择。扫描件、缺页、读不清公式都应报告，不猜。

专项许可scope：transform:图ID、add-image:图ID、redraw:图ID、supplement:内容ID、font-change。每项提案文件写明源文件、区域/内容、理由和用途。plan.permissions引用scope和proposal路径；生成前比对确认指纹。批准完整原图不自动批准裁剪，持有原始数据不等于允许重绘。

记录命令不能核实回复是谁说的，哈希只能证明文件未变，不能证明阅读正确或用户批准真实。代理对真实性、完整精读和原文忠实负责。

## 初次开发本包

已获制作skill授权。实际版式代表页先提交选用内容与3—5张页面清单，用户确认后才做。初版先展示可检查结果、验证范围与建议版本号，按确认流程再合入稳定版。
