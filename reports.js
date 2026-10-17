/* Local-only handwritten forms and workbook previews. No network or patient seed. */
(function(root){
 'use strict';
 const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
 const col=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
 const num=s=>[...s].reduce((v,c)=>v*26+c.charCodeAt(0)-64,0);
 const attr=(s,k)=>s.match(new RegExp('(?:^|\\s)'+k+'="([^"]*)"'))?.[1]||'';
 const roles={nurse:'看護',rehab:'リハ',care:'介護'};
 function monthBefore(date){const [y,m]=date.split('-').map(Number);return `${m===1?y-1:y}-${String(m===1?12:m-1).padStart(2,'0')}`;}
 function validateOptions(o){if(!Number.isInteger(o.year)||o.year<2000||o.year>2100||![1,4,7,10].includes(o.startMonth)||!/^\d{4}-(0[1-9]|1[0-2])$/.test(o.referenceMonth))throw Error('対象年・開始月・参照月を確認してください。');}
 function sorted(s){return s.residents.filter(p=>p.active).slice().sort((a,b)=>a.bed.localeCompare(b.bed,undefined,{numeric:true}));}
 function pages(title,headers,rows,widths,subtitle){
  const sheets=[];for(let i=0;i<Math.max(1,rows.length);i+=20)sheets.push({name:`${title}${Math.floor(i/20)+1}`,title,subtitle,headers,rows:rows.slice(i,i+20),widths});return sheets;
 }
 function build(kind,s,o){
  const people=sorted(s);
  if(kind==='weight'){
   validateOptions(o);const term=[1,4].includes(o.startMonth)?'前期':'後期';
   const headers=['部屋・ベッド','氏名',`${o.referenceMonth.replace('-','年')}月\n体重 (kg)`,...Array.from({length:6},(_,i)=>(o.startMonth+i-1)%12+1+'月')];
   return pages('体重記入表',headers,people.map(p=>[p.bed,p.name,p.weightReferences?.[o.referenceMonth]??'',...Array(6).fill('')]),[12,23,16,...Array(6).fill(14)],`${o.year}年 ${term}（${o.startMonth}月〜${(o.startMonth+4)%12+1}月）／参照値は指定月のみ。空欄は未入力。`);
  }
  if(kind==='patient')return pages('患者別担当表',['部屋・ベッド','氏名','看護','リハ','介護'],people.map(p=>[p.bed,p.name,...Object.keys(roles).map(k=>p.assigned[k]||'未設定')]),[14,24,24,24,24],`作成日 ${o.date}／現在の担当割`);
  if(kind==='staff'){
   const rows=[];
   for(const [role,label] of Object.entries(roles)){
    // Include legacy assignments whose staff member is absent from the active roster.
    const names=[...new Set([...s.staff.filter(x=>x.active&&x.role===role).map(x=>x.name),...people.map(p=>p.assigned[role]).filter(Boolean)])];
    for(const name of [...names,'']){const assigned=people.filter(p=>(p.assigned[role]||'')===name);if(!name&&!assigned.length)continue;
     if(!assigned.length)rows.push([label,name,'','担当患者なし']);
     for(const p of assigned)rows.push([label,name||'未設定',p.bed,p.name]);
    }
   }
   return pages('スタッフ別担当表',['職種','スタッフ','部屋・ベッド','担当患者'],rows,[12,28,16,32],`作成日 ${o.date}／職種別・スタッフ名簿順。未設定も掲載。`);
  }
  throw Error('帳票の種類を選んでください。');
 }
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 function workbook(sheets){
  const entries=new Map(),lastCol=s=>col(s.headers.length);
  entries.set('[Content_Types].xml',`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`);
  entries.set('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  entries.set('xl/workbook.xml',`<workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s,i)=>`<sheet name="${esc(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets><definedNames>${sheets.map((s,i)=>`<definedName name="_xlnm.Print_Area" localSheetId="${i}">'${esc(s.name)}'!$A$1:$${lastCol(s)}$${s.rows.length+3}</definedName>`).join('')}</definedNames></workbook>`);
  entries.set('xl/_rels/workbook.xml.rels',`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  const xf=(font,border,numFmt=0)=>`<xf numFmtId="${numFmt}" fontId="${font}" fillId="0" borderId="${border}" xfId="0" applyAlignment="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>`;
  entries.set('xl/styles.xml',`<styleSheet xmlns="${ns}"><fonts count="2"><font><sz val="11"/><name val="Yu Gothic"/></font><font><b/><sz val="15"/><name val="Yu Gothic"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="2"><border/><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4">${xf(0,0)}${xf(1,0)}${xf(0,1)}${xf(0,1,2)}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
  sheets.forEach((s,i)=>{
   const cell=(v,r,c,style)=>`<c r="${col(c)}${r}" s="${style}"${typeof v==='number'?'':' t="inlineStr"'}>${typeof v==='number'?`<v>${v}</v>`:`<is><t xml:space="preserve">${esc(v)}</t></is>`}</c>`;
   const rows=[[s.title],[s.subtitle],s.headers,...s.rows];
   entries.set(`xl/worksheets/sheet${i+1}.xml`,`<worksheet xmlns="${ns}"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:${lastCol(s)}${rows.length}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${s.widths.map((w,j)=>`<col min="${j+1}" max="${j+1}" width="${w}" customWidth="1"/>`).join('')}</cols><sheetData>${rows.map((row,r)=>`<row r="${r+1}" ht="${r===0?30:r===1?24:r===2?36:26}" customHeight="1">${row.map((v,c)=>cell(v,r+1,c+1,r===0?1:r<2?0:typeof v==='number'?3:2)).join('')}</row>`).join('')}</sheetData><mergeCells count="2"><mergeCell ref="A1:${lastCol(s)}1"/><mergeCell ref="A2:${lastCol(s)}2"/></mergeCells><printOptions horizontalCentered="1"/><pageMargins left="0.25" right="0.25" top="0.3" bottom="0.3" header="0.1" footer="0.1"/><pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="1"/><headerFooter><oddFooter>&amp;C${i+1} / ${sheets.length}</oddFooter></headerFooter></worksheet>`);
  });
  return root.WardExcel.zip(entries);
 }
 function reportHTML(sheets){return sheets.map((s,i)=>`<section class="report-paper"><div class="report-page-number">${i+1} / ${sheets.length}</div><h2>${esc(s.title)}</h2><p>${esc(s.subtitle)}</p><table><colgroup>${s.widths.map(w=>`<col style="width:${w}ch">`).join('')}</colgroup><thead><tr>${s.headers.map(v=>`<th>${esc(v)}</th>`).join('')}</tr></thead><tbody>${s.rows.map(row=>`<tr>${row.map(v=>`<td>${esc(typeof v==='number'?v.toFixed(2):v)}</td>`).join('')}</tr>`).join('')}</tbody></table></section>`).join('');}
 // Read the generated OOXML, including merged cells, column widths and cached values.
 // Excel-specific drawing objects and conditional formatting are outside this preview.
 function excelHTML(w){
  const parse=xml=>new DOMParser().parseFromString(xml,'application/xml');
  const styles=parse(w.text('xl/styles.xml')),fonts=[...styles.querySelectorAll('fonts > font')],fills=[...styles.querySelectorAll('fills > fill')],xfs=[...styles.querySelectorAll('cellXfs > xf')];
  const rgb=e=>{const value=e?.getAttribute('rgb');return /^[0-9A-Fa-f]{8}$/.test(value||'')?'#'+value.slice(2):'';};
  const css=style=>{const xf=xfs[style],font=fonts[Number(xf?.getAttribute('fontId')||0)],fill=fills[Number(xf?.getAttribute('fillId')||0)],align=xf?.querySelector('alignment');return `font-size:${Number(font?.querySelector('sz')?.getAttribute('val')||11)}pt;${font?.querySelector('b')?'font-weight:bold;':''}${rgb(font?.querySelector('color'))?'color:'+rgb(font.querySelector('color'))+';':''}${rgb(fill?.querySelector('fgColor'))?'background:'+rgb(fill.querySelector('fgColor'))+';':''}text-align:${['center','right','left'].includes(align?.getAttribute('horizontal'))?align.getAttribute('horizontal'):'left'};`;};
  const names=parse(w.text('xl/workbook.xml')),areas=new Map();
  for(const e of names.querySelectorAll('definedName[name="_xlnm.Print_Area"]'))areas.set(Number(e.getAttribute('localSheetId')),e.textContent);
  return Object.entries(w.sheets).map(([name,s],index)=>{
   if(!areas.has(index))return '';const doc=parse(s.xml),area=areas.get(index).match(/\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)/);if(!area)return '';
   const [,a,b,c,d]=area,startCol=num(a),endCol=num(c),startRow=Number(b),endRow=Number(d);if(endCol>100||endRow>500)throw Error('プレビュー範囲が大きすぎます。');
   const merges=new Map(),covered=new Set();for(const e of doc.querySelectorAll('mergeCell')){const [from,to]=e.getAttribute('ref').split(':'),fc=num(from.match(/[A-Z]+/)[0]),fr=Number(from.match(/\d+/)[0]),tc=num(to.match(/[A-Z]+/)[0]),tr=Number(to.match(/\d+/)[0]);merges.set(from,{cols:tc-fc+1,rows:tr-fr+1});for(let r=fr;r<=tr;r++)for(let c=fc;c<=tc;c++)if(r!==fr||c!==fc)covered.add(col(c)+r);}
   const hidden=new Set(),widths={};for(const e of doc.querySelectorAll('cols > col'))for(let c=Number(e.getAttribute('min'));c<=Number(e.getAttribute('max'));c++){if(e.getAttribute('hidden')==='1')hidden.add(c);widths[c]=Number(e.getAttribute('width')||10);}
   const rowInfo=new Map([...doc.querySelectorAll('sheetData > row')].map(e=>[Number(e.getAttribute('r')),e]));
   let html=`<section class="report-paper"><h2>${esc(name)}</h2><table class="excel-preview"><colgroup>`;for(let c=startCol;c<=endCol;c++)if(!hidden.has(c))html+=`<col style="width:${widths[c]||10}ch">`;html+='</colgroup><tbody>';
   for(let r=startRow;r<=endRow;r++){const info=rowInfo.get(r);if(info?.getAttribute('hidden')==='1')continue;html+=`<tr style="height:${Number(info?.getAttribute('ht')||20)}pt">`;for(let c=startCol;c<=endCol;c++){const ref=col(c)+r;if(hidden.has(c)||covered.has(ref))continue;const v=s.cells[ref]||{value:'',style:0},m=merges.get(ref);let value=v.value;const fmt=Number(xfs[v.style]?.getAttribute('numFmtId')||0);if(typeof value==='number'&&[14,15,16,17,22].includes(fmt)){const date=new Date(Date.UTC(1899,11,30)+value*86400000);value=`${date.getUTCMonth()+1}/${date.getUTCDate()}`;}
    html+=`<td ${m?`colspan="${m.cols}" rowspan="${m.rows}"`:''} style="${css(v.style)}">${esc(value)}</td>`;}html+='</tr>';}
   return html+'</tbody></table></section>';
  }).join('');
 }
 root.WardReports={build,workbook,reportHTML,excelHTML,monthBefore,validateOptions};
})(globalThis);
