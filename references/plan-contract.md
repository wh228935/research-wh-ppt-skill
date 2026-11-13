# 文件约定与制作入口

所有路径相对于它所在的JSON文件解析（project.json中的确认保存绝对路径）。工作目录在skill包外。对象尺寸以PowerPoint point为单位，960×540为默认16:9，不用像素混入字号。

## intake.json

```json
{"duration_minutes":15,"background":"用户实际说明","documents":[{"id":"paper1","path":"paper.pdf","sha256":"实际文件哈希","chapters":[{"id":"c1","title":"原文正文第1章"}]}]}
```

这只是字段示意，不是用户已确定15分钟或文档阅读已完成。按实际文件登记所有正文章节；附录按用户选择纳入。sha256用真实文件计算。

## outline.json（用户确认的内容，不随小排版改变）

```json
{"pages":[{"id":"s01","title":"当页主题","header":"1. 章节名","on_screen":["完整上屏句子","图题","表格每格文字"],"assets":["fig1"],"layout_suggestion":"图与解释的关系","speaking_points":["本页讲解重点"],"estimated_seconds":60,"decisions":[],"backup":false}]}
```

上屏文字包括实际计划出现的表格单元格、图注和必要标签。允许排版换行、坐标和样式改变；不允许未确认的新句子。结构页也计入大纲。详细备注来自实际精读，不一定逐字在大纲列出，但不能增加未支持结论。

## figures.json / prototype-selection.json

图件选择文件在真实图表清单、缩略图和用户选择基础上记录`{"selected_assets":["fig1"],"files":[{"path":"media/original.png","sha256":"真实图件哈希"}]}`；清单全文作为工作证据保留。files绑定所选原图的实际版本。无图用空列表，不捏造选图。

代表页选择用`{"pages":["s01","s04","s08"]}`，3—5个不同ID。仅示意，不预定真实任务页序。

## plan.json（渲染计划）

```json
{
 "outline_file":"outline.json","style_file":"style.json",
 "assets":[{"id":"fig1","path":"media/original.png","sha256":"真实文件哈希","required_permissions":[]}],
 "permissions":[],
 "pages":[{"id":"s01","title":"当页主题","header":"1. 章节名","notes":"真实详细讲稿",
 "objects":[{"id":"figure1","type":"image","asset_id":"fig1","box":[24,120,580,350],"fit":"contain"},
 {"id":"interpretation","type":"text","text":"完整上屏句子","box":[630,130,300,250],"font_size":18,"color":"243447","bold":false}]}]
}
```

示例坐标/字号仅说明字段，代表页后才能成为本任务批准风格。复制assets/style.json到任务目录调整，不改全局包。style中header_mode为section或topic，可在page.header_mode按已确认的页面方案覆盖；section另画当页主题，正文对象应避开标题区。不要再把页标题重复放入objects。逐页页眉方式应在代表页审核与最终风格记录中写清，不能生成全稿时自行改用另一种。

对象：text（text/font_size/color/bold/align/highlights）、rect（fill/line/rounded）、image（asset_id，只contain，不裁）、table（rows等宽二维数组/font_size）、line（box为起点x/y与终点差值w/h；from/to连接已有对象，默认矩形右侧4接左侧2，竖向可指定from_site/to_site；arrow为true）。每个对象id唯一、box为4个数。表题单独text对象放表上，图题放图下。正文默认左对齐。高亮项`{"text":"准确词句","color":"FF0000"}`仅改变样式，不改文本。

asset.origin为source（默认）、external、processed或redrawn；source/external文件须被figures.files绑定。processed/redrawn另写original_path及original_sha256对应已选原图，强制transform:ID或redraw:ID；external强制add-image:ID。required_permissions可附加其他scope。plan.permissions列`{"scope":"transform:fig1","proposal":"proposals/fig1.json"}`。处理范围和输出须忠于批准提案，脚本只能核对文件与声明，不能识别图意。实际处理另行执行并保留原图，不能把fit改cover绕过许可。字体替换需要font-change提案。

`Build-Deck.ps1 -PlanPath ... -WorkDir ... -OutputPath ... -Stage prototype|full -Python ...`只创建新稿，不自动重建局改原文件。生成前检查确认记录；渲染计划文字/图件须匹配内容大纲。微小排版不改outline文件，实质内容变更先重新确认outline。

prototypes确认针对审核快照，`project.py snapshot --output prototype-review.json prototypes.pptx style.json previews/render-report.json`；快照files绑定真实文件。用户看后确认，`record --scope prototypes ...`登记。全稿使用同一风格文件。

## bundle.json

```json
{"pptx":"final.pptx","docx":"notes.docx","preview_dir":"previews-powerpoint","contact_sheet":"contact-sheet.png","qa":"qa.json","software_reports":{"PowerPoint":"previews-powerpoint/render-report.json","WPS":"previews-wps/render-report.json"},"formal_pages":15,"backup_pages":2}
```

qa.json记真实内容检查、逐页目视检查：source_sha256、content_reviewed、visual_reviewed、unresolved_errors。报告不是自动推断，没检查不得填true。按实际页数填写，数字只是字段示意。

逐页图和缩略图由Render-Office.ps1与Contact-Sheet.ps1产生；Word由office.py notes-docx从最终PPT生成。归档：`project.py archive --workdir ... --plan ... --manifest bundle.json --output <成品根目录> --revision v01`。已有v01不可覆盖；内部draft临时文件不占新稿号。
