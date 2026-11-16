// Editable reference layouts, CSS px: 1280x720 = 960x540 pt. No runtime imports.
export const palette={blue:'#005BA6',panel:'#145DA5',pale:'#D6E5F1',band:'#DFEFF9',ink:'#193B5D',gray:'#536A7F',red:'#F00000'};
export function box(s,name,x,y,w,h,fill,o={}){return s.shapes.add({name,geometry:o.rounded?'roundRect':'rect',position:{left:x,top:y,width:w,height:h},fill,line:{fill:o.line||'none',width:o.line?1:0,style:'solid'},...(o.rounded?{borderRadius:o.rounded}:{}),...(o.shadow?{shadow:'shadow-md'}:{})});}
export function text(s,name,value,x,y,w,h,size=28,o={}){const t=box(s,name,x,y,w,h,'none');t.text=value;t.text.style={typeface:'Microsoft YaHei',fontSize:size,color:o.color||palette.ink,bold:o.bold||false,alignment:o.align||'left',verticalAlignment:'middle',insets:0,autoFit:'none',wrap:'square'};return t;}
export function header(s,title){box(s,'ReferenceHeader',0,0,1280,86,palette.blue);text(s,'PageTitle',title,48,15,1184,56,36,{bold:true,color:'#FFFFFF'});}
export function conclusion(s,value){box(s,'ConclusionBand',64,633,1152,71,palette.band,{rounded:15,shadow:true});box(s,'ConclusionLabel',64,633,93,71,palette.blue,{rounded:15});text(s,'ConclusionLabelText','结论',74,642,73,51,29,{color:'#FFFFFF',align:'center'});text(s,'ConclusionChevron','»',172,644,26,49,35,{color:'#0077BE'});return text(s,'ConclusionText',value,204,640,990,57,25,{color:'#172F46'});}
export function addImage(s,a,[x,y,w,h]){if(!(a.width_px>0&&a.height_px>0&&a.blob))throw Error('Read original image bytes and actual dimensions before placing a figure');const k=Math.min(w/a.width_px,h/a.height_px),fw=a.width_px*k,fh=a.height_px*k;return s.images.add({blob:a.blob,contentType:a.contentType||'image/png',alt:a.alt||'原图',fit:'contain',position:{left:x+(w-fw)/2,top:y+(h-fh)/2,width:fw,height:fh}});}
export function applyLayout(s,role,d={}){
 s.background.fill='#FFFFFF';
 if(role==='cover'){
  s.shapes.add({name:'CoverBlueCurve',geometry:'custom',position:{left:0,top:0,width:1280,height:475},fill:palette.blue,line:{fill:'none',width:0},customPaths:[{width:1280,height:475,commands:[{moveTo:{x:0,y:0}},{lineTo:{x:1280,y:0}},{lineTo:{x:1280,y:420}},{cubicBezTo:{x1:920,y1:492,x2:360,y2:492,x:0,y:420}},{close:{}}]}]});
  text(s,'ReportType',d.report_type||'汇报类型',72,70,1136,44,27,{color:'#FFFFFF',align:'center'});
  text(s,'PageTitle',d.title||'汇报题目',72,177,1136,135,50,{color:'#FFFFFF',align:'center'});
  if(d.subtitle)text(s,'EnglishTitle',d.subtitle,100,325,1080,82,27,{color:'#FFFFFF',align:'center'});
  text(s,'CoverAuthors',d.authors||'按本次任务确认的汇报信息填写',110,515,1060,52,24,{color:palette.blue});
  text(s,'CoverVenue',d.venue||'单位、场景或文献出处',110,579,1060,45,27,{color:palette.blue,bold:true});
  if(d.detail)text(s,'CoverDetail',d.detail,110,641,1060,35,21,{color:palette.gray});
 }else if(role==='contents'){
  box(s,'ContentsBluePanel',0,0,478,720,palette.blue,{shadow:true});
  text(s,'PageTitle','目 录',30,282,418,70,62,{color:'#FFFFFF',bold:true,align:'center'});
  text(s,'ContentsEnglish','CONTENTS',30,354,418,64,57,{color:'#FFFFFF',bold:true,align:'center'});
  const list=d.chapters||['第一部分','第二部分','第三部分','第四部分'],n=list.length;
  if(n<1||n>6)throw Error('Propose another contents layout above 6 entries');
  const h=n>5?59:70,g=n===1?0:Math.min(74,(624-n*h)/(n-1)),y0=(720-n*h-(n-1)*g)/2;
  list.forEach((v,i)=>{const y=y0+i*(h+g);box(s,`ContentsNumber${i+1}`,643,y,70,h,palette.blue,{rounded:13,shadow:true});box(s,`ContentsRow${i+1}`,757,y,370,h,palette.blue,{rounded:13,shadow:true});text(s,`ContentsNumberText${i+1}`,String(i+1),649,y+5,58,h-10,34,{color:'#FFFFFF',align:'center'});text(s,`ContentsTitle${i+1}`,v,774,y+6,336,h-12,29,{color:'#FFFFFF',bold:true,align:'center'});});
 }else if(role==='body-figure'){
  header(s,d.header||'章节编号与当页主题');
  if(d.title)text(s,'BodyTopic',d.title,42,103,1196,54,32,{bold:true,color:palette.blue});
  if(d.image)addImage(s,d.image,[42,169,1196,344]);
  else{box(s,'FigureSlot',65,181,1150,318,'#F7F9FB',{line:'#CADDEC'});text(s,'FigurePlaceholder','原图放置区域：完整保留比例、坐标和图例',95,287,1090,64,27,{color:palette.gray,align:'center'});}
  const a=d.details||['按内容放置图题、数值比较或说明'];a.forEach((v,i)=>text(s,`FigureDetail${i+1}`,v,45+i*(1190/a.length),540,1170/a.length,70,24,{color:i===0?palette.blue:palette.ink,bold:i===0}));
  if(d.conclusion!==false)conclusion(s,d.conclusion||'依据本次材料确认的本页小结');
 }else if(role==='summary'){
  header(s,d.header||'章节编号 结论');
  const a=d.items||[{label:'主要结论一',text:'依据原文填写结论，保留必要的条件与范围。'},{label:'主要结论二',text:'依据真实内容调整行数和行高，不为固定行数补写结论。'},{label:'主要结论三',text:'内容放不下时先建议拆页，不缩小文字硬塞。'}];
  if(a.length<1||a.length>5)throw Error('Propose splitting the summary above 5 rows');
  const top=137,bottom=d.conclusion===false?675:601,gap=24,h=Math.min(133,(bottom-top-gap*(a.length-1))/a.length),start=top+(bottom-top-a.length*h-gap*(a.length-1))/2;
  a.forEach((v,i)=>{const y=start+i*(h+gap);box(s,`SummaryLabel${i+1}`,65,y,322,h,palette.panel,{rounded:8});box(s,`SummaryBody${i+1}`,405,y,824,h,palette.pale,{rounded:8});text(s,`SummaryNumber${i+1}`,String(i+1).padStart(2,'0'),84,y+12,48,h-24,28,{color:'#FFFFFF',bold:true});text(s,`SummaryTitle${i+1}`,v.label,154,y+12,220,h-24,28,{color:'#FFFFFF',bold:true});text(s,`SummaryText${i+1}`,v.text,428,y+12,778,h-24,27);});
  if(d.conclusion!==false)conclusion(s,d.conclusion||'根据以上结论概括本次研究，不增加未经确认的新判断');
 }else if(role==='closing'){
  box(s,'ReferenceHeader',0,0,1280,86,palette.blue);
  box(s,'ClosingOuter',16,100,1240,596,'#FFFFFF',{rounded:78,line:'#D3E4F2'});
  box(s,'ClosingInner',110,150,1060,466,'#F5F8FB',{rounded:64,line:'#D3E4F2'});
  text(s,'PageTitle',d.title||'感谢各位老师',170,255,940,74,51,{color:palette.blue,bold:true,align:'center'});
  text(s,'ClosingSubtitle',d.subtitle||'恳请批评指正',170,374,940,60,34,{align:'center'});
  text(s,'ClosingTopic',d.topic||'本次汇报题目',170,558,940,39,22,{color:palette.gray,align:'center'});
 }else throw Error('Unknown reference layout: '+role);
 return s;
}
