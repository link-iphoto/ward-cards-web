/* Small OOXML patcher: untouched ZIP entries, conditional formats, printer setup,
   shapes and extensions are retained. No workbook conversion library is used. */
(function(root){
 'use strict';
 const enc=new TextEncoder(),dec=new TextDecoder();
 const esc=s=>String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
 const unesc=s=>String(s??'').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&');
 const attr=(s,k)=>unesc(s.match(new RegExp('(?:^|\\s)'+k+'="([^"]*)"'))?.[1]||'');
 const crcTable=Array.from({length:256},(_,i)=>{for(let j=0;j<8;j++)i=(i&1)?0xEDB88320^(i>>>1):i>>>1;return i>>>0;});
 function crc(bytes){let c=0xffffffff;for(const b of bytes)c=crcTable[(c^b)&255]^(c>>>8);return (c^0xffffffff)>>>0;}
 async function unzip(bytes) {
  const v=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);let end=-1;
  for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(v.getUint32(i,true)===0x06054b50){end=i;break;}
  if(end<0)throw Error('Excelファイルを読み取れません。xlsxを選んでください。');
  const entries=new Map();let p=v.getUint32(end+16,true),total=0;
  for(let n=0;n<v.getUint16(end+10,true);n++){
   if(v.getUint32(p,true)!==0x02014b50)throw Error('ZIP構造が不正です。');
   const method=v.getUint16(p+10,true),size=v.getUint32(p+20,true),rawSize=v.getUint32(p+24,true),nl=v.getUint16(p+28,true),xl=v.getUint16(p+30,true),cl=v.getUint16(p+32,true),off=v.getUint32(p+42,true);
   total+=rawSize;if(total>80*1024*1024)throw Error('Excelが大きすぎます。');
   const name=dec.decode(bytes.slice(p+46,p+46+nl)),start=off+30+v.getUint16(off+26,true)+v.getUint16(off+28,true);
   let data=bytes.slice(start,start+size);
   if(method===8)data=new Uint8Array(await new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());
   else if(method!==0)throw Error('対応していないExcel圧縮形式です。');
   if(data.length!==rawSize || crc(data)!==v.getUint32(p+16,true))throw Error('Excelファイルの内容が破損しています。');
   entries.set(name,data);p+=46+nl+xl+cl;
  }
  return entries;
 }
 function zip(entries) {
  let offset=0;const local=[],central=[];let centralSize=0;
  for(const [name,raw] of entries){
   const data=typeof raw==='string'?enc.encode(raw):raw,n=enc.encode(name),sum=crc(data);
   const h=new Uint8Array(30+n.length),v=new DataView(h.buffer);v.setUint32(0,0x04034b50,true);v.setUint16(4,20,true);v.setUint16(6,0x800,true);v.setUint32(14,sum,true);v.setUint32(18,data.length,true);v.setUint32(22,data.length,true);v.setUint16(26,n.length,true);h.set(n,30);local.push(h,data);
   const c=new Uint8Array(46+n.length),cv=new DataView(c.buffer);cv.setUint32(0,0x02014b50,true);cv.setUint16(4,20,true);cv.setUint16(6,20,true);cv.setUint16(8,0x800,true);cv.setUint32(16,sum,true);cv.setUint32(20,data.length,true);cv.setUint32(24,data.length,true);cv.setUint16(28,n.length,true);cv.setUint32(42,offset,true);c.set(n,46);central.push(c);centralSize+=c.length;offset+=h.length+data.length;
  }
  const end=new Uint8Array(22),ev=new DataView(end.buffer);ev.setUint32(0,0x06054b50,true);ev.setUint16(8,entries.size,true);ev.setUint16(10,entries.size,true);ev.setUint32(12,centralSize,true);ev.setUint32(16,offset,true);
  const out=new Uint8Array(offset+centralSize+22);let pos=0;for(const part of [...local,...central,end]){out.set(part,pos);pos+=part.length;}return out;
 }
 function parseCells(xml,strings=[]) {
  const out={};for(const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)){
   const a=m[1],body=m[2]||'',t=attr(a,'t'),v=unesc(body.match(/<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/)?.[1]||'');
   let value=t==='s'?(strings[Number(v)]??''):t==='inlineStr'?[...body.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(x=>unesc(x[1])).join(''):t==='str'?v:v===''?'':Number(v);
   out[attr(a,'r')]={value,formula:unesc(body.match(/<f(?:\s[^>]*)?>([\s\S]*?)<\/f>/)?.[1]||''),style:Number(attr(a,'s')||0)};
  }return out;
 }
 async function read(bytes) {
  const entries=await unzip(bytes),text=n=>dec.decode(entries.get(n)||new Uint8Array()),strings=[...text('xl/sharedStrings.xml').matchAll(/<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g)].map(m=>[...m[1].replace(/<rPh\b[^>]*>[\s\S]*?<\/rPh>/g,'').matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(t=>unesc(t[1])).join(''));
  const rel={};for(const m of text('xl/_rels/workbook.xml.rels').matchAll(/<Relationship\b([^>]*)\/>/g))rel[attr(m[1],'Id')]=attr(m[1],'Target');
  const sheets={};for(const m of text('xl/workbook.xml').matchAll(/<sheet\b([^>]*)\/>/g)){
   const target=rel[attr(m[1],'r:id')];if(!target)continue;const path=target.startsWith('/')?target.slice(1):'xl/'+target;
   sheets[attr(m[1],'name')]={path,xml:text(path),cells:parseCells(text(path),strings),patches:{}};
  }
  return {entries,sheets,text,strings};
 }
 function set(w,s,ref,value,formula,style){if(!w.sheets[s])throw Error('シートが見つかりません：'+s);w.sheets[s].patches[ref]={value,formula,style};}
 function get(w,s,ref){return w.sheets[s].patches[ref]?.value??w.sheets[s].cells[ref]?.value??'';}
 const colNumber=ref=>[...ref.match(/^[A-Z]+/)[0]].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);
 function patchSheet(sheet) {
  const pending={...sheet.patches};
  function cell(ref,oldAttrs=''){
   const p=pending[ref];delete pending[ref];let a=oldAttrs.replace(/\s*(?:r|t|s)="[^"]*"/g,'').trim();const st=p.style??sheet.cells[ref]?.style;
   const attributes=` r="${ref}"${st!==undefined?` s="${st}"`:''}${a?' '+a:''}`;
   if(p.formula!==undefined){const t=typeof p.value==='number'?'':' t="str"';return `<c${attributes}${t}><f>${esc(p.formula.replace(/^=/,''))}</f><v>${esc(p.value)}</v></c>`;}
   if(p.value===''||p.value===null||p.value===undefined)return `<c${attributes}/>`;
   return typeof p.value==='number'?`<c${attributes}><v>${p.value}</v></c>`:`<c${attributes} t="inlineStr"><is><t xml:space="preserve">${esc(p.value)}</t></is></c>`;
  }
  let xml=sheet.xml.replace(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g,(all,attrs,body='')=>{
   const row=Number(attr(attrs,'r'));body=body.replace(/<c\b([^>]*?)(?:\/>|>[\s\S]*?<\/c>)/g,(old,a)=>{const r=attr(a,'r');return pending[r]?cell(r,a):old;});
   const extra=Object.keys(pending).filter(r=>Number(r.match(/\d+$/)[0])===row);
   if(extra.length){const parts=[...body.matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)].map(m=>m[0]);for(const r of extra)parts.push(cell(r));parts.sort((a,b)=>colNumber(attr(a,'r'))-colNumber(attr(b,'r')));body=parts.join('');}
   return `<row${attrs}>${body}</row>`;
  });
  if(Object.keys(pending).length)throw Error('原本にない行への出力を防止しました：'+Object.keys(pending).join(','));
  return xml;
 }
 function styleFills(w) {
  let xml=w.text('xl/styles.xml');const m=xml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);const styles=[...m[1].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map(x=>x[0]);
  const fills=xml.match(/<fills\b[^>]*>([\s\S]*?)<\/fills>/),items=[...fills[1].matchAll(/<fill\b[^>]*>[\s\S]*?<\/fill>/g)].map(x=>x[0]),gray=items.length;
  items.push('<fill><patternFill patternType="solid"><fgColor rgb="FF888888"/><bgColor indexed="64"/></patternFill></fill>');xml=xml.replace(fills[0],`<fills count="${items.length}">${items.join('')}</fills>`);const cache={};
  return {style(base,grey){const key=base+':'+grey;if(cache[key]!==undefined)return cache[key];let s=styles[base].replace(/fillId="\d+"/,`fillId="${grey?gray:0}"`);if(!s.includes('applyFill='))s=s.replace('<xf ','<xf applyFill="1" ');else s=s.replace(/applyFill="[^"]*"/,'applyFill="1"');cache[key]=styles.length;styles.push(s);return cache[key];},finish(){xml=xml.replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/,`<cellXfs count="${styles.length}">${styles.join('')}</cellXfs>`);w.entries.set('xl/styles.xml',enc.encode(xml));}};
 }
 const bytes64=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
 function nameStyles(w,sourceSheet){
  let xml=w.text('xl/styles.xml');const fonts=[...xml.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/)[1].matchAll(/<font\b[^>]*>[\s\S]*?<\/font>/g)].map(m=>m[0]);
  const styles=[...xml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map(m=>m[0]);
  const borders=[...xml.match(/<borders\b[^>]*>([\s\S]*?)<\/borders>/)[1].matchAll(/<border\b[^>]*>[\s\S]*?<\/border>/g)].map(m=>m[0]);
  const sourceColor=ref=>fonts[Number(attr(styles[w.sheets[sourceSheet].cells[ref].style],'fontId'))].match(/<color\b[^>]*\/>/)?.[0];
  const colors={red:sourceColor('A3')||'<color rgb="FFFF0000"/>',blue:sourceColor('C3')||'<color rgb="FF365F91"/>',black:'<color rgb="FF000000"/>',gray:'<color theme="0" tint="-0.499984740745262"/>'};
  const cache={},fontCache={};
  return {style(base,color,strong=undefined,edge=''){const key=[base,color,strong,edge].join(':');if(cache[key]!==undefined)return cache[key];const xf=styles[base],fontId=Number(attr(xf,'fontId')),fk=[fontId,color,strong].join(':');let fid=fontCache[fk];
   if(fid===undefined){let f=fonts[fontId];f=/<color\b/.test(f)?f.replace(/<color\b[^>]*\/>/,colors[color]||colors.black):f.replace('</font>',(colors[color]||colors.black)+'</font>');if(strong!==undefined)f=f.replace(/<b(?:\s[^>]*)?\/>/g,'');if(strong)f=f.replace('</font>','<b/></font>');fid=fonts.length;fonts.push(f);fontCache[fk]=fid;}
   let s=xf.replace(/fontId="\d+"/,`fontId="${fid}"`);if(!s.includes('applyFont='))s=s.replace('<xf ','<xf applyFont="1" ');else s=s.replace(/applyFont="[^"]*"/,'applyFont="1"');
   if(edge){const width=strong?'medium':'thin',border=`<border><left style="${width}"/><right style="${width}"/><top${edge==='top'?` style="${width}"`:''}/><bottom${edge==='bottom'?` style="${width}"`:''}/><diagonal/></border>`,bid=borders.length;borders.push(border);s=s.replace(/borderId="\d+"/,`borderId="${bid}"`);s=s.includes('applyBorder=')?s.replace(/applyBorder="[^"]*"/,'applyBorder="1"'):s.replace('<xf ','<xf applyBorder="1" ');}
   cache[key]=styles.length;styles.push(s);return cache[key];
  },finish(){xml=xml.replace(/<fonts\b[^>]*>[\s\S]*?<\/fonts>/,`<fonts count="${fonts.length}">${fonts.join('')}</fonts>`).replace(/<borders\b[^>]*>[\s\S]*?<\/borders>/,`<borders count="${borders.length}">${borders.join('')}</borders>`).replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/,`<cellXfs count="${styles.length}">${styles.join('')}</cellXfs>`);w.entries.set('xl/styles.xml',enc.encode(xml));}};
 }
 function bathBorders(w){
  let xml=w.text('xl/styles.xml');const styles=[...xml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map(m=>m[0]);const cache={};
  for(const sn of ['月木am','月木pm','火金am','火金pm'])for(let r=3;r<=24;r++)for(const col of 'ABCDEFGHIJ'){
   const p=w.sheets[sn].patches[col+r],base=p.style??w.sheets[sn].cells[col+r]?.style??0,border=attr(styles[w.sheets['月木am'].cells[col+'3'].style],'borderId'),key=base+':'+border;
   if(cache[key]===undefined){let s=styles[base].replace(/borderId="\d+"/,`borderId="${border}"`);s=s.includes('applyBorder=')?s.replace(/applyBorder="[^"]*"/,'applyBorder="1"'):s.replace('<xf ','<xf applyBorder="1" ');cache[key]=styles.length;styles.push(s);}p.style=cache[key];
  }
  xml=xml.replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/,`<cellXfs count="${styles.length}">${styles.join('')}</cellXfs>`);w.entries.set('xl/styles.xml',enc.encode(xml));
 }
 function careRoomBorders(w,ends){
  let xml=w.text('xl/styles.xml');const styles=[...xml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)[1].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map(m=>m[0]);
  const borders=[...xml.match(/<borders\b[^>]*>([\s\S]*?)<\/borders>/)[1].matchAll(/<border\b[^>]*>[\s\S]*?<\/border>/g)].map(m=>m[0]),cache={};const sheet=w.sheets['排泄ケアチェック'];
  for(const [cols,rows] of ends)for(const r of rows)for(const col of cols){const ref=col+r,old=sheet.patches[ref],base=old?.style??sheet.cells[ref]?.style??0,key=base+':double';
   if(cache[key]===undefined){let border=borders[Number(attr(styles[base],'borderId'))].replace(/<bottom\b[^>]*(?:\/>|>[\s\S]*?<\/bottom>)/,'<bottom style="double"/>');const bid=borders.length;borders.push(border);let st=styles[base].replace(/borderId="\d+"/,`borderId="${bid}"`);st=st.includes('applyBorder=')?st.replace(/applyBorder="[^"]*"/,'applyBorder="1"'):st.replace('<xf ','<xf applyBorder="1" ');cache[key]=styles.length;styles.push(st);}
   sheet.patches[ref]={...old,value:old?.value??sheet.cells[ref]?.value??'',formula:(old?.formula??sheet.cells[ref]?.formula)||undefined,style:cache[key]};
  }
  xml=xml.replace(/<borders\b[^>]*>[\s\S]*?<\/borders>/,`<borders count="${borders.length}">${borders.join('')}</borders>`).replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/,`<cellXfs count="${styles.length}">${styles.join('')}</cellXfs>`);w.entries.set('xl/styles.xml',enc.encode(xml));
 }
 function hideHelper(sheet,col){const n=colNumber(col+'1'),def=`<col min="${n}" max="${n}" width="3" hidden="1" customWidth="1"/>`;sheet.xml=sheet.xml.includes('</cols>')?sheet.xml.replace('</cols>',def+'</cols>'):sheet.xml.replace(/(<sheetFormatPr\b[^>]*\/>)/,'$1<cols>'+def+'</cols>');}
 function dateSerial(date){return Math.round((Date.parse(date+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000);}
 // Insert detail rows in the embedded original, keeping native extensions intact.
 // Only the six bath sheets are expanded; the user's source files are untouched.
 function insertDetailRows(w,sn,at,count) {
  const s=w.sheets[sn],sample=s.xml.match(new RegExp('<row\\b[^>]*\\br="'+(at-1)+'"[^>]*>[\\s\\S]*?<\\/row>'))?.[0];
  if(!sample)throw Error('増設元の明細行が見つかりません：'+sn);
  const shiftRefs=text=>text.replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g,(all,col,row)=>col+(Number(row)>=at?Number(row)+count:row));
  let xml=s.xml.replace(/\b(ref|sqref|activeCell|topLeftCell)="([^"]*)"/g,(all,key,value)=>`${key}="${shiftRefs(value)}"`)
   .replace(/(<row\b[^>]*\br=")(\d+)(")/g,(all,a,n,b)=>a+(Number(n)>=at?Number(n)+count:n)+b)
   .replace(/(<c\b[^>]*\br=")([A-Z]+)(\d+)(")/g,(all,a,col,n,b)=>a+col+(Number(n)>=at?Number(n)+count:n)+b)
   .replace(/(<brk\b[^>]*\bid=")(\d+)(")/g,(all,a,n,b)=>a+(Number(n)>=at?Number(n)+count:n)+b);
  // Existing own-sheet footer and conditional format references shift as in Excel.
  xml=xml.replace(/<(f|formula)\b(?![^>]*\/>)([^>]*)>([\s\S]*?)<\/\1>/g,(all,tag,a,body)=>`<${tag}${a}>${body.includes('!')?body:shiftRefs(body)}</${tag}>`);
  let rows='';for(let i=0;i<count;i++)rows+=sample.replace(/(<row\b[^>]*\br=")\d+"/,'$1'+(at+i)+'"').replace(/<c\b([^>]*?)(?:\/>|>[\s\S]*?<\/c>)/g,(_,attrs)=>`<c${attrs.replace(/\br="([A-Z]+)\d+"/,(_,col)=>`r="${col}${at+i}"`).replace(/\s+t="[^"]*"/g,'')}/>`);
  xml=xml.replace(new RegExp('(<row\\b[^>]*\\br="'+(at+count)+'")'),rows+'$1');
  const merged=[...s.xml.matchAll(/<mergeCell\b[^>]*ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"[^>]*\/>/g)].filter(m=>Number(m[2])===at-1&&m[2]===m[4]);
  if(merged.length){let extra='';for(let i=0;i<count;i++)for(const m of merged)extra+=`<mergeCell ref="${m[1]}${at+i}:${m[3]}${at+i}"/>`;xml=xml.replace('</mergeCells>',extra+'</mergeCells>').replace(/<mergeCells count="(\d+)"/,(_,n)=>`<mergeCells count="${Number(n)+count*merged.length}"`);}
  s.xml=xml;s.cells=parseCells(xml,w.strings);
 }
 function nextWeekday(date,days){let d=new Date(date+'T12:00:00');for(let i=0;i<7;i++){if(days.includes(d.getDay()))return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;d.setDate(d.getDate()+1);}throw Error('日付が不正です');}
 async function generate(key,state,seed,date) {
  state=WardCore.syncRoomGroups(WardCore.copy(state));
  const w=await read(bytes64(seed.templates[key]));const people=new Map(state.residents.filter(p=>p.active).map(p=>[p.id,p])),bedMap=new Map([...people.values()].map(p=>[p.bed,p]));
  const list=k=>(state.orders[k]||[]).map(id=>people.get(id)).filter(Boolean),bathSlots=k=>(state.orders[k]||[]).map(id=>people.get(id)||null),put=(s,r,v,f,st)=>set(w,s,r,v,f,st);
  const bind=(s,r,p,field)=>{if(p&&w.sheets[s].patches[r])w.sheets[s].patches[r].binding={id:p.id,field};};
  if(key==='map'||key==='food'){
   const ns=nameStyles(w,key==='map'?'２階':'元データ');
   const colorOf=p=>p?.nameColor||seed.state.residents.find(x=>x.id===p?.id)?.nameColor||'black';
   const s=key==='map'?'２階':'元データ';for(const slot of seed.slots){const p=bedMap.get(slot.bed);put(s,slot.cell,p?.name||'');bind(s,slot.cell,p,'name');if(key==='food'){put(s,slot.mealCell,p?.meal||'');bind(s,slot.mealCell,p,'meal');}}
   for(const slot of seed.slots){const p=bedMap.get(slot.bed),color=colorOf(p);w.sheets[s].patches[slot.cell].style=ns.style(w.sheets[s].cells[slot.cell]?.style||0,color);if(key==='food')w.sheets[s].patches[slot.mealCell].style=ns.style(w.sheets[s].cells[slot.mealCell]?.style||0,color);}
   if(key==='map'){put(s,'D21',dateSerial(date),'TODAY()');put(s,'C23',`　　　${people.size}床（入所中）`);}
   else {
    for(const sn of ['談話','Ａ','Ｂ'])for(const [r,c] of Object.entries(w.sheets[sn].cells))if(c.formula){const refs=[...c.formula.matchAll(/元データ!\$?([A-Z]+)\$?(\d+)/g)];if(!refs.length)throw Error('食事用紙の式を確認してください：'+sn+'!'+r);put(sn,r,get(w,'元データ',refs[0][1]+refs[0][2]),c.formula);}
    for(const [r,c] of Object.entries(w.sheets['経管'].cells))if(c.formula){
     if(r==='A21'){put('経管',r,'経管','"経管"');continue;}
     const refs=[...c.formula.matchAll(/元データ!\$?([A-Z]+)\$?(\d+)/g)];if(!refs.length)throw Error('経管の式を確認してください：'+r);
     const sr=refs[0][1]+refs[0][2],v=get(w,'元データ',sr),next=get(w,'元データ',refs[0][1]+(Number(refs[0][2])+1));
     put('経管',r,typeof v==='number'?v:next==='経管'&&!['談話','A','B','経管','禁食'].includes(v)?v:'',c.formula);
    }
    const slotByCell=new Map(seed.slots.flatMap(slot=>[[slot.cell,slot],[slot.mealCell,slot]]));
    for(const sn of ['談話','Ａ','Ｂ','経管'])for(const [ref,patch] of Object.entries(w.sheets[sn].patches)){
     const source=patch.formula?.match(/元データ!\$?([A-Z]+)\$?(\d+)/),slot=source?slotByCell.get(source[1]+source[2]):null;if(!slot)continue;
     const p=bedMap.get(slot.bed),category={談話:'談話',Ａ:'A',Ｂ:'B',経管:'経管'}[sn];
     const matched=p?.meal===category,isName=source[1]+source[2]===slot.cell;
     patch.style=ns.style(w.sheets[sn].cells[ref]?.style||0,matched?colorOf(p):'gray',matched,sn==='経管'?'':isName?'top':'bottom');
    }
   }
   ns.finish();
  }
  if(key==='laundry')for(const [group,sn] of [['laundry-A','Ａグループ'],['laundry-B','Ｂグループ']]){
   const arr=list(group);for(let i=0;i<30;i++){const p=arr[i],r=i+5,room=p?Number(p.bed.split('-')[0]):'',same=p&&arr[i-1]&&p.bed.split('-')[0]===arr[i-1].bed.split('-')[0];put(sn,'T'+r,room);put(sn,'A'+r,room&&!same?room:'',`IF(T${r}="","",IF(T${r}=T${r-1},"",T${r}))`);put(sn,'B'+r,p?.name||'');bind(sn,'B'+r,p,'name');put(sn,'C'+r,p?'様':'',`IF(B${r}="","","様")`);}
   hideHelper(w.sheets[sn],'T');
  }
  if(key==='bath'){
   for(const sn of ['月木am','月木pm','火金am','火金pm'])insertDetailRows(w,sn,sn==='火金pm'?23:22,sn==='火金pm'?1:2);
   for(const sn of ['月木洗濯表','火金洗濯表']){insertDetailRows(w,sn,24,2);insertDetailRows(w,sn,54,2);}
   for(const [group,sn] of [['bath-mon-am','月木am'],['bath-mon-pm','月木pm'],['bath-tue-am','火金am'],['bath-tue-pm','火金pm']]){
    const arr=bathSlots(group),day=nextWeekday(date,group.includes('mon')?[1,4]:[2,5]);if(arr.length>22)throw Error('入浴は各枠22行までです。');put(sn,'E1',dateSerial(day),group.endsWith('-pm')?`${group.includes('mon')?'月木':'火金'}am!E1`:undefined);put(sn,'H1',arr.filter(Boolean).length,'COUNTA(B3:B24)');
    for(let i=0;i<22;i++){
     const p=arr[i],r=i+3,b=p?.bath||{},room=p?Number(p.bed.split('-')[0]):'',previous=arr[i-1],same=previous&&room===Number(previous.bed.split('-')[0]);
     put(sn,'M'+r,room);put(sn,'A'+r,room&&!same?room:'',`IF(M${r}="","",IF(M${r}=M${r-1},"",M${r}))`);
     const short=v=>({'ストレッチャー':'スト','車椅子':'車'})[v]||v||'';
     for(const [col,val] of Object.entries({B:p?.name||'',D:b.sheets||'',E:b.kind||'',F:'',G:'',H:short(b.outbound),I:short(b.change),J:short(b.inbound)}))put(sn,col+r,val);
     for(const [col,field] of Object.entries({B:'name',D:'bath.sheets',E:'bath.kind',H:'bath.outbound',I:'bath.change',J:'bath.inbound'}))bind(sn,col+r,p,field);
     put(sn,'C'+r,p?'様':'',`IF(B${r}="","","様")`);
    }
    hideHelper(w.sheets[sn],'M');
   }
   for(const [prefix,base] of [['月木','mon'],['火金','tue']]){
    const sn=prefix+'洗濯表',day=nextWeekday(date,base==='mon'?[1,4]:[2,5]);put(sn,'G1',dateSerial(day),`${prefix}am!E1`);put(sn,'G31',dateSerial(day),'G1');
    for(const [suffix,start] of [['am',5],['pm',35]])for(let i=0;i<22;i++){
     const p=bathSlots(`bath-${base}-${suffix}`)[i],r=start+i,sr=i+3,arr=bathSlots(`bath-${base}-${suffix}`),same=p&&arr[i-1]&&p.bed.split('-')[0]===arr[i-1].bed.split('-')[0];
     put(sn,'A'+r,p&&!same?Number(p.bed.split('-')[0]):'',`IF(${prefix}${suffix}!A${sr}="","",${prefix}${suffix}!A${sr})`);
     put(sn,'B'+r,p?.name||'',`IF(${prefix}${suffix}!B${sr}="","",${prefix}${suffix}!B${sr})`);
     put(sn,'C'+r,p?'様':'',`IF(B${r}="","","様")`);
    }
   }
   bathBorders(w);
  }
  if(key==='care'){
   const rowMap=new Map();for(let i=0;i<seed.beds.length;i++){
    const bed=seed.beds[i],r=i+2,p=bedMap.get(bed);if(p)rowMap.set(p.id,r);
    for(const [col,v] of Object.entries({A:bed,B:p?.name||'',C:p?.assigned.nurse||'',D:p?.assigned.rehab||'',E:p?.assigned.care||''}))put('患者一覧',col+r,v);
    for(const [col,field] of Object.entries({B:'name',C:'assigned.nurse',D:'assigned.rehab',E:'assigned.care'}))bind('患者一覧',col+r,p,field);
   }
   for(const [role,col,countCol] of [['nurse','A','B'],['care','D','E'],['rehab','G','H']]){
    const names=state.staff.filter(p=>p.active&&p.role===role);if(names.length>15)throw Error('スタッフ一覧は職種ごと15人までです。');
    for(let i=0;i<15;i++){const name=names[i]?.name||'',r=i+2,ref=countCol+r;put('リスト',col+r,name);put('リスト',ref,name?[...people.values()].filter(p=>p.assigned[role]===name).length:0,w.sheets['リスト'].cells[ref]?.formula||`COUNTIF(テーブル1[${{nurse:'看護',care:'介護',rehab:'リハ'}[role]}],リスト!${col}${r})`);}
   }
   const sf=styleFills(w),sn='排泄ケアチェック';
   for(const [group,cols,cap] of [['care-A',['A','B','C','E','F','G'],28],['care-B',['I','J','K','M','N','O'],30]]){
    const side=group.slice(-1),beds=seed.beds.filter(b=>WardCore.roomGroup({bed:b})===side).sort((a,b)=>a.localeCompare(b,undefined,{numeric:true})),rooms=[...new Set(beds.map(b=>b.split('-')[0]))],arr=[],ends=[];
    const placed=state.orders[group]||[];
    if(placed.length>cap)throw Error(group+'の配置がベッド総数を超えています。');
    const plans=rooms.map(room=>{const roomBeds=beds.filter(b=>b.startsWith(room+'-')),members=placed.filter(id=>WardCore.careBlankRoom(id)===room||people.get(id)?.bed.startsWith(room+'-')),vacant=roomBeds.filter(b=>!bedMap.has(b));return {room,members,vacant,count:Math.max(0,roomBeds.length-members.length)};});
    let surplus=plans.reduce((n,x)=>n+x.count,0)-(cap-placed.length);
    for(let i=plans.length-1;i>=0&&surplus;i--){const trim=Math.min(surplus,plans[i].count);plans[i].count-=trim;surplus-=trim;}
    for(const plan of plans){for(const id of plan.members){const p=people.get(id);arr.push({p,bed:p?.bed||''});}for(const bed of plan.vacant.slice(0,plan.count))arr.push({p:null,bed});ends.push(arr.length+10);}
    if(arr.length!==cap)throw Error(group+'のベッド枠と用紙の行数が一致しません。');
    for(let i=0;i<cap;i++){
     const {p,bed}=arr[i],r=i+11,source=p?rowMap.get(p.id):!bedMap.has(bed)?seed.beds.indexOf(bed)+2:null,flags=p?WardCore.careCells(p):{remark:'',checks:['','',''],gray:true};
     put(sn,cols[0]+r,p?.bed||'',source?`IF(患者一覧!B${source}="","",患者一覧!A${source})`:undefined);
     put(sn,cols[1]+r,p?.name||'',source?`IF(患者一覧!B${source}="","",患者一覧!B${source})`:undefined);
     bind(sn,cols[0]+r,p,'bed');bind(sn,cols[1]+r,p,'name');
     put(sn,cols[2]+r,flags.remark);
     for(let j=0;j<3;j++){const ref=cols[j+3]+r;put(sn,ref,flags.checks[j],undefined,sf.style(w.sheets[sn].cells[ref]?.style||0,j===1&&flags.gray));}
    }
    if(!w.roomEnds)w.roomEnds=[];w.roomEnds.push([side==='A'?'ABCDEFG':'IJKLMNO',ends]);
   }
   put(sn,'A1','（　　／　　）排泄ケアチェック');sf.finish();careRoomBorders(w,w.roomEnds);
   // These existing auxiliary sheets remain linked to the canonical patient table.
   for(const sn of ['体重A','体重B'])for(const [r,c] of Object.entries(w.sheets[sn].cells))if(c.formula){const ref=c.formula.match(/患者一覧!([AB]\d+)/);if(ref)put(sn,r,get(w,'患者一覧',ref[1]),c.formula);}
  }
  const printRanges={food:{'談話':'$A$1:$G$34','Ａ':'$A$1:$E$35','Ｂ':'$A$1:$F$42','経管':'$A$1:$G$29'},map:{'２階':'$A$1:$G$29'},care:{'排泄ケアチェック':'$A$1:$O$40'},laundry:{'Ａグループ':'$A$1:$S$38','Ｂグループ':'$A$1:$S$38'},bath:{'月木am':'$A$1:$J$25','月木pm':'$A$1:$J$25','火金am':'$A$1:$J$25','火金pm':'$A$1:$J$25','月木洗濯表':'$A$1:$R$60','火金洗濯表':'$A$1:$R$60'}}[key];
  for(const sn of Object.keys(printRanges)){
   const s=w.sheets[sn],height=sn.endsWith('洗濯表')?2:1;
   if(/<pageSetUpPr\b/.test(s.xml))s.xml=s.xml.replace(/<pageSetUpPr\b[^>]*\/>/,'<pageSetUpPr fitToPage="1"/>');
   else if(/<sheetPr>/.test(s.xml))s.xml=s.xml.replace('</sheetPr>','<pageSetUpPr fitToPage="1"/></sheetPr>');
   else if(/<sheetPr\s[^>]*>/.test(s.xml))s.xml=s.xml.replace('</sheetPr>','<pageSetUpPr fitToPage="1"/></sheetPr>');
   else s.xml=s.xml.replace(/(<worksheet\b[^>]*>)/,'$1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
   s.xml=s.xml.replace(/<pageSetup\b([^>]*)\/>/,(all,a)=>`<pageSetup${a.replace(/\s(?:fitToHeight|fitToWidth|scale)="[^"]*"/g,'')} fitToWidth="1" fitToHeight="${height}"/>`);
  }
  const snapshot={};for(const [sn,s] of Object.entries(w.sheets))if(Object.keys(s.patches).length){w.entries.set(s.path,enc.encode(patchSheet(s)));snapshot[sn]=s.patches;}
  let wb=w.text('xl/workbook.xml');wb=wb.replace(/<calcPr\b[^>]*\/>/,'<calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1" calcMode="auto"/>');
  let newNames='';const allSheets=Object.keys(w.sheets);
  for(const [sn,range] of Object.entries(printRanges)){
   const id=allSheets.indexOf(sn);wb=wb.replace(new RegExp('<definedName\\b(?=[^>]*name="_xlnm.Print_Area")(?=[^>]*localSheetId="'+id+'")[^>]*>[\\s\\S]*?<\\/definedName>','g'),'');
   newNames+=`<definedName name="_xlnm.Print_Area" localSheetId="${id}">'${esc(sn)}'!${range}</definedName>`;
  }
  wb=wb.includes('</definedNames>')?wb.replace('</definedNames>',newNames+'</definedNames>'):wb.replace('</sheets>','</sheets><definedNames>'+newNames+'</definedNames>');
  w.entries.set('xl/workbook.xml',enc.encode(wb));
  if(w.entries.has('xl/calcChain.xml')){
   w.entries.delete('xl/calcChain.xml');w.entries.set('xl/_rels/workbook.xml.rels',enc.encode(w.text('xl/_rels/workbook.xml.rels').replace(/<Relationship\b[^>]*Type="[^"]*\/calcChain"[^>]*\/>/g,'')));
   w.entries.set('[Content_Types].xml',enc.encode(w.text('[Content_Types].xml').replace(/<Override\b[^>]*PartName="\/xl\/calcChain.xml"[^>]*\/>/g,'')));
  }
  return {bytes:zip(w.entries),snapshot};
 }
 root.WardExcel={zip,unzip,read,generate,parseCells,bytes64,dateSerial};
})(globalThis);
