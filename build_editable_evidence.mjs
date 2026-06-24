import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const OUT = "F:/Desk/面试资料/ai-agent-interview-guide-main/project-python/outputs/多模态证据矩阵_可编辑.pptx";
const TMP = "C:/Users/cbc24/AppData/Local/Temp/codex-presentations/manual-20260623/editable-evidence-diagram";
const C = { navy:"#123F68", navy2:"#0D355B", line:"#6F7478", border:"#7D858B", pale:"#E9F3FB",
  pale2:"#F4F9FC", coral:"#F26D5B", teal:"#20A6A1", blue:"#3F86DC", yellow:"#F4B21A",
  ink:"#20262D", white:"#FFFFFF" };

const deck = Presentation.create({ slideSize: { width: 1280, height: 907 } });
const slide = deck.slides.add();
slide.background.fill = C.white;

function shape(geometry,name,x,y,w,h,fill="none",lineFill="none",lineWidth=0,extra={}) {
  return slide.shapes.add({ geometry, name, position:{left:x,top:y,width:w,height:h,...(extra.rotation!==undefined?{rotation:extra.rotation}:{})},
    fill, line:{style:extra.lineStyle||"solid",fill:lineFill,width:lineWidth},
    ...(extra.borderRadius?{borderRadius:extra.borderRadius}:{}) });
}
function text(name,value,x,y,w,h,size=16,color=C.ink,bold=false,align="left",valign="middle",extra={}) {
  const s=shape("textbox",name,x,y,w,h,extra.fill||"none",extra.lineFill||"none",extra.lineWidth||0,{borderRadius:extra.borderRadius});
  s.text=value;
  s.text.style={fontSize:size,color,bold,alignment:align,verticalAlignment:valign,autoFit:"shrinkText",wrap:"square",
    insets:extra.insets||{top:2,right:4,bottom:2,left:4},typeface:extra.typeface||"Microsoft YaHei"};
  return s;
}
function circle(name,label,x,y,d,fill,size=18,color=C.white) {
  const s=shape("ellipse",name,x,y,d,d,fill,"none",0);
  if(label){s.text=label;s.text.style={fontSize:size,color,bold:true,alignment:"center",verticalAlignment:"middle",autoFit:"shrinkText",
    insets:{top:0,right:0,bottom:0,left:0},typeface:"Microsoft YaHei"};}
  return s;
}
function line(name,x1,y1,x2,y2,width=2,color=C.line,dashed=false) {
  return shape("line",name,Math.min(x1,x2),Math.min(y1,y2),Math.max(1,Math.abs(x2-x1)),Math.max(1,Math.abs(y2-y1)),"none",color,width,{lineStyle:dashed?"dashed":"solid"});
}
function triangle(name,x,y,size,rotation=90,fill=C.line){return shape("triangle",name,x,y,size,size,fill,"none",0,{rotation});}
function arrowRight(name,x1,y,x2,color=C.line,width=2){line(name+"-line",x1,y,x2-10,y,width,color);triangle(name+"-head",x2-13,y-6,12,90,color);}
function arrowDown(name,x,y1,y2,color=C.line,width=2){line(name+"-line",x,y1,x,y2-10,width,color);triangle(name+"-head",x-6,y2-13,12,180,color);}
function elbowRight(name,x1,y1,elbowX,y2,x2,color=C.line,width=2){
  line(name+"-h1",x1,y1,elbowX,y1,width,color);line(name+"-v",elbowX,y1,elbowX,y2,width,color);
  line(name+"-h2",elbowX,y2,x2-10,y2,width,color);triangle(name+"-head",x2-13,y2-6,12,90,color);
}
function container(name,x,y,w,h){return shape("roundRect",name,x,y,w,h,"none",C.border,1.4,{borderRadius:18,lineStyle:"dashed"});}
function evidenceIcon(name,x,y,scale=1){
  const w=48*scale,h=55*scale;
  shape("roundRect",name+"-back2",x+12*scale,y-10*scale,w,h,C.pale2,"#B7C4CE",1,{borderRadius:7});
  shape("roundRect",name+"-back1",x+6*scale,y-5*scale,w,h,C.pale2,"#9EAFBC",1,{borderRadius:7});
  shape("roundRect",name+"-front",x,y,w,h,C.white,"#7E96A8",1.2,{borderRadius:7});
  for(let i=0;i<3;i++){circle(name+"-dot"+i,"",x+8*scale,y+(12+i*13)*scale,5*scale,C.blue);
    line(name+"-row"+i,x+17*scale,y+(14+i*13)*scale,x+39*scale,y+(14+i*13)*scale,2*scale,C.blue);}
}
function llmIcon(name,x,y,d){
  circle(name+"-bg","",x,y,d,C.navy);const cx=x+d/2,cy=y+d/2;
  const core=circle(name+"-core","",cx-d*.10,cy-d*.10,d*.20,"none");core.line={style:"solid",fill:C.white,width:2.4};
  for(let i=0;i<6;i++){const a=i*Math.PI/3,ex=cx+Math.cos(a)*d*.22-d*.12,ey=cy+Math.sin(a)*d*.22-d*.17;
    shape("ellipse",name+"-petal"+i,ex,ey,d*.24,d*.34,"none",C.white,2.2,{rotation:i*60});}
}
function dotGrid(name,x,y,cols,rows,gapX,gapY,d,fill){for(let r=0;r<rows;r++)for(let c=0;c<cols;c++)circle(`${name}-${r}-${c}`,"",x+c*gapX,y+r*gapY,d,fill);}

// Connectors first.
arrowDown("user-to-request",90,155,205); elbowRight("request-to-parser",145,280,80,510,101);
arrowRight("parser-llm-to-slots",187,555,225); arrowRight("slots-to-decision",335,555,392);
line("decision-to-trunk",438,555,455,555,2,C.line); line("branch-trunk",455,133,455,555,2,C.line);
arrowRight("branch-b",455,133,482); arrowRight("branch-c",455,343,482); arrowRight("branch-d",455,540,482);
arrowRight("b-r-to-oq",590,130,705); arrowRight("b-oq-to-evidence",775,130,830);
arrowRight("c-r-to-grid",565,338,650); arrowRight("c-grid-to-teal",715,338,785); arrowRight("c-teal-to-rank",830,338,860);
arrowRight("c-rank-to-evidence",900,338,935); arrowRight("d-rdo-to-candidates",565,545,655);
arrowRight("d-candidates-to-filter",735,545,790); arrowRight("d-filter-to-topk",830,545,855); arrowRight("d-topk-to-evidence",890,545,935);
line("aggregate-trunk",955,130,955,550,2,C.line); line("aggregate-b",900,130,955,130,2,C.line);
arrowRight("aggregate-to-matrix",955,390,990); arrowDown("matrix-to-generation",1118,530,652);
arrowRight("prompt-to-llm",760,760,815); arrowRight("llm-to-response",905,760,955);

// Main containers and query parsing.
container("parser-container",100,405,275,300); container("mode-b-container",482,18,450,220);
container("mode-c-container",482,250,450,205); container("mode-d-container",482,468,450,165);
container("generation-container",480,650,750,215);

shape("roundRect","user-prompt-bubble",57,85,68,52,C.white,C.navy,3,{borderRadius:10});
const tail=triangle("user-prompt-tail",80,126,18,180,C.white); tail.line={style:"solid",fill:C.navy,width:2.4};
circle("bubble-dot-1","",77,107,7,C.navy); circle("bubble-dot-2","",91,107,7,C.navy); circle("bubble-dot-3","",105,107,7,C.navy);
text("user-prompt-label","user prompt",45,140,94,28,17,C.ink,false,"center");
text("request-box","描述患者症状！\n调用医学知识图谱",18,205,155,95,16,C.white,true,"center","middle",
  {fill:C.navy,borderRadius:14,insets:{top:8,right:8,bottom:8,left:8}});

text("parser-title","(a) LLM query parsing",112,420,235,32,20,C.ink,true);
llmIcon("parser-llm",120,500,68); text("parser-llm-label","LLM",119,570,70,30,17,C.ink,true,"center");
shape("roundRect","slot-box",225,485,112,145,C.white,C.navy,2,{borderRadius:14});
text("slot-r","r   疾病",237,497,88,28,17,C.ink,true); text("slot-d","d   症状",237,527,88,28,17,C.ink,true);
text("slot-m","m  影像",237,557,88,28,17,C.ink,true); text("slot-o","o₍q₎  概念",237,587,88,28,17,C.ink,true);
const decision=shape("diamond","condition-decision",392,515,72,72,C.navy,"none",0);
decision.text="分层\n判定"; decision.text.style={fontSize:14,color:C.white,bold:true,alignment:"center",verticalAlignment:"middle",autoFit:"shrinkText",
  insets:{top:0,right:0,bottom:0,left:0},typeface:"Microsoft YaHei"};

// (b) Known disease: concept verification.
text("mode-b-title","(b) 已知疾病 · 概念验证（r + o₍q₎）",500,28,400,34,20,C.ink,true);
circle("b-r","r",525,96,58,C.coral,20); text("b-rdo","RDO",620,88,60,26,17,C.ink,false,"center");
circle("b-oq","o₍q₎",710,96,58,C.teal,18); evidenceIcon("b-evidence",835,93,.85);
text("b-evidence-label","证据",832,153,72,28,16,C.ink,true,"center");
text("b-caption","定位单条 RDO 边，查找对应证据",540,176,335,30,16,C.ink,true,"center");

// (c) Known disease: dimension retrieval.
text("mode-c-title","(c) 已知疾病 · 维度检索（o₍q₎ 为空）",500,258,400,34,20,C.ink,true);
circle("c-r","r",505,313,58,C.coral,20); text("c-rdu","RDU",585,300,50,25,17,C.ink,false,"center");
text("c-rdu-note","（疾病与症状）",565,358,95,24,14,C.ink,false,"center");
dotGrid("c-blue-grid",664,302,2,4,28,28,15,C.blue);
text("c-ruo","Rᵤₒ",733,300,45,25,17,C.ink,false,"center");
text("c-ruo-note","（顶点分组聚合）",713,358,115,24,14,C.ink,false,"center");
dotGrid("c-teal-stack",795,303,1,3,0,34,21,C.teal);
for(let i=0;i<4;i++)text("c-rank-"+i,String(i+1),855,295+i*24,28,24,14,C.ink,true,"center","middle",{fill:C.white,lineFill:C.border,lineWidth:1});
text("c-rank-label","f（疾病特征）\n（模态维度）",875,300,58,62,12,C.ink,true,"center","middle",{insets:{top:0,right:0,bottom:0,left:0}});
evidenceIcon("c-evidence",935,315,.55); text("c-evidence-label","证据",926,375,62,26,16,C.ink,true,"center");

// (d) Unknown disease: reverse symptom retrieval.
text("mode-d-title","(d) 未知疾病 · 症状反查（无 r）",500,474,400,32,20,C.ink,true);
circle("d-oq-plus","o₁",500,510,32,C.teal,13); text("d-plus-1","+",536,510,22,26,18,C.ink,true,"center");
circle("d-oq-minus","o₂",500,547,32,C.yellow,13,C.navy); text("d-minus","−",536,547,22,26,20,C.ink,true,"center");
circle("d-oq-plus2","o₃",500,584,32,C.blue,13); text("d-plus-2","+",536,584,22,26,18,C.ink,true,"center");
text("d-oq-n","…   o₍q₎ × n",495,612,98,22,14,C.ink,false);
text("d-rdo-reverse","RDO 反向",582,528,73,25,15,C.ink,false,"center");
dotGrid("d-candidates",670,515,3,4,28,24,14,C.coral); text("d-candidates-label","候选疾病 × N",650,607,105,22,14,C.ink,true,"center");
shape("funnel","d-filter",790,512,52,70,C.navy,C.navy2,1);
text("d-filter-label","筛选排序 ·\n聚合、去重证据",760,590,112,36,13,C.ink,true,"center");
dotGrid("d-top-k",868,515,1,4,0,24,14,C.coral); text("d-top-k-label","Top-K",848,607,65,22,15,C.ink,true,"center");
evidenceIcon("d-evidence",935,520,.55); text("d-evidence-label","证据",926,578,62,26,16,C.ink,true,"center");

// Multimodal evidence matrix; every cell and marker remains editable.
const mx=990,my=244,colW=[90,49,49,49,49],rowH=[72,48,61,61,61];
text("matrix-title","多模态信息矩阵\n(Multimodal Evidence Matrix)",mx,my,286,rowH[0],17,C.white,true,"center","middle",
  {fill:C.navy,borderRadius:12,insets:{top:6,right:4,bottom:4,left:4}});
const headers=["模态","文本","结构化","影像","诊疗"], rows=["临床表现","检查","治疗"];
const dots=[[C.coral,C.teal,C.blue,C.yellow],[C.blue,C.teal,C.coral,C.yellow],[C.yellow,C.teal,C.blue,C.coral]];
let cy=my+rowH[0],cx=mx;
for(let c=0;c<5;c++){text("matrix-header-"+c,headers[c],cx,cy,colW[c],rowH[1],15,C.ink,true,"center","middle",
  {fill:C.white,lineFill:"#87939C",lineWidth:1.1});cx+=colW[c];}
cy+=rowH[1];
for(let r=0;r<3;r++){cx=mx;text("matrix-row-label-"+r,rows[r],cx,cy,colW[0],rowH[r+2],16,C.ink,true,"center","middle",
  {fill:C.white,lineFill:"#87939C",lineWidth:1.1});cx+=colW[0];
  for(let c=0;c<4;c++){shape("rect",`matrix-cell-${r}-${c}`,cx,cy,colW[c+1],rowH[r+2],C.white,"#87939C",1.1);
    circle(`matrix-dot-${r}-${c}`,"",cx+colW[c+1]/2-7,cy+rowH[r+2]/2-7,14,dots[r][c]);cx+=colW[c+1];}
  cy+=rowH[r+2];
}
text("aggregate-label","Augment",956,398,34,18,10,C.ink,false,"center","middle",{insets:{top:0,right:0,bottom:0,left:0}});

// (e) Generation.
text("generation-title","(e) generation",500,660,300,32,20,C.ink,true);
shape("roundRect","input-prompt-frame",520,705,245,125,C.white,C.navy,2,{borderRadius:12});
text("input-prompt-title","Input prompts",535,708,215,30,18,C.ink,true,"center");
text("input-matrix","多模态信息矩阵",545,744,195,36,16,C.white,true,"center","middle",{fill:C.navy,borderRadius:8});
text("input-user","user prompt",545,786,195,32,16,C.navy,false,"center","middle",
  {fill:C.pale,lineFill:"#7E96A8",lineWidth:1.2,borderRadius:8});
llmIcon("generation-llm",820,718,90); text("generation-llm-label","LLM",830,810,70,28,17,C.ink,true,"center");
text("response-box","Response:\n候选疾病与可追溯证据",960,710,240,110,20,C.white,true,"center","middle",
  {fill:C.navy,borderRadius:12,insets:{top:10,right:10,bottom:10,left:10}});

await fs.mkdir(path.dirname(OUT),{recursive:true}); await fs.mkdir(path.join(TMP,"preview"),{recursive:true});
const preview=await deck.export({slide,format:"png",scale:1});
await fs.writeFile(path.join(TMP,"preview","slide-01.png"),new Uint8Array(await preview.arrayBuffer()));
const layout=await slide.export({format:"layout"});
await fs.writeFile(path.join(TMP,"preview","slide-01.layout.json"),await layout.text());
const pptx=await PresentationFile.exportPptx(deck); await pptx.save(OUT); console.log(OUT);
