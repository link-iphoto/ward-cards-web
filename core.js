(function (root) {
  'use strict';
  const GROUPS = {
    'bath-mon-am': {label:'月・木 午前', family:'bath', cap:22},
    'bath-mon-pm': {label:'月・木 午後', family:'bath', cap:22},
    'bath-tue-am': {label:'火・金 午前', family:'bath', cap:22},
    'bath-tue-pm': {label:'火・金 午後', family:'bath', cap:22},
    'care-A': {label:'Aグループ', family:'care', cap:28},
    'care-B': {label:'Bグループ', family:'care', cap:30},
    'laundry-A': {label:'Aグループ', family:'laundry', cap:30},
    'laundry-B': {label:'Bグループ', family:'laundry', cap:30}
  };
  const copy = x => JSON.parse(JSON.stringify(x));
  const isBathBlank = id => typeof id==='string' && id.startsWith('blank:bath:');
  const isCareBlank = id => typeof id==='string' && /^blank:care:\d+:/.test(id);
  const careBlankRoom = id => isCareBlank(id)?id.split(':')[2]:'';
  function roomGroup(p) {const n=Number(String(p.bed||'').split('-')[0]);return n>=215&&n<=222?'A':(n>=202&&n<=213)||[223,225].includes(n)?'B':'';}
  // Keep manual ordering within each side. Only room transfers cross A/B.
  function syncRoomGroups(s) {
    for(const family of ['care','laundry']) {
      const moved={A:[],B:[]};
      for(const side of ['A','B'])s.orders[family+'-'+side]=s.orders[family+'-'+side].filter(id=>{
        const p=s.residents.find(p=>p.id===id),target=p&&roomGroup(p);
        if(target&&target!==side){moved[target].push(id);return false;}return true;
      });
      for(const side of ['A','B'])s.orders[family+'-'+side].push(...moved[side]);
    }
    const byId=id=>s.residents.find(p=>p.id===id);
    for(const p of s.residents)if(p.bath?.kind==='ストレッチャー')p.bath.kind='ミスト';
    for(const side of ['A','B']){
      s.orders['care-'+side].sort((a,b)=>(careBlankRoom(a)||room(byId(a))).localeCompare(careBlankRoom(b)||room(byId(b)),undefined,{numeric:true}));
      const pending=s.orders['laundry-'+side].filter(id=>id.startsWith('?'));
      s.orders['laundry-'+side]=s.residents.filter(p=>eligible(p,'laundry')&&roomGroup(p)===side).sort((a,b)=>a.bed.localeCompare(b.bed,undefined,{numeric:true})).map(p=>p.id).concat(pending);
    }
    return s;
  }
  function room(p){return p?String(p.bed||'').split('-')[0]:'未確認';}
  function eligible(p,family) {return p.active && (family==='bath'?p.bathEligible:family==='care'?p.careEligible:true);}
  function move(state,id,to,index) {
    const p=state.residents.find(p=>p.id===id); const family=GROUPS[to]?.family || to.replace('unassigned-','');
    const blank=isBathBlank(id)||isCareBlank(id), targetRoom=isCareBlank(id)?careBlankRoom(id):room(p);
    if ((!p && !blank) || (isBathBlank(id) && family!=='bath') || (isCareBlank(id) && family!=='care') || !['bath','care','laundry'].includes(family) || (p && !eligible(p,family))) throw Error('この配置の対象者ではありません。');
    if(isCareBlank(id)&&!GROUPS[to])throw Error('排泄の空白行は部屋の中だけで移動できます。');
    if(GROUPS[to]&&family==='care'&&!to.endsWith('-'+roomGroup({bed:targetRoom+'-1'})))throw Error('この部屋は別のグループです。');
    if(GROUPS[to]&&family==='laundry'&&!to.endsWith('-'+roomGroup(p)))throw Error('この部屋は別のグループです。');
    if(family==='laundry')throw Error('洗濯表は部屋・ベッド番号の昇順固定です。個人カードでベッドを変更してください。');
    if(family==='care'&&GROUPS[to]){
      const a=state.orders[to],indices=a.map((x,i)=>(careBlankRoom(x)||room(state.residents.find(p=>p.id===x)))===targetRoom?i:-1).filter(i=>i>=0);
      if(indices.length&&(index<indices[0]||index>indices[indices.length-1]+1))throw Error('排泄チェックは同じ部屋の中だけで順番を変更できます。');
    }
    const before=copy(state.orders);
    const old=(state.orders[to]||[]).indexOf(id);
    for (const [key,g] of Object.entries(GROUPS)) if(g.family===family) state.orders[key]=state.orders[key].filter(x=>x!==id);
    if (GROUPS[to]) {
      if(old>=0 && old<index) index--;
      const a=state.orders[to];a.splice(Math.max(0,Math.min(index,a.length)),0,id);
      if(a.length>GROUPS[to].cap){state.orders=before;throw Error(`${GROUPS[to].label}は用紙の上限${GROUPS[to].cap}人です。別の枠へ移動してください。`);}
    }
  }
  function moveStaff(state,id,role,index) {
    const staff=state.staff.find(s=>s.id===id&&s.active);
    if(!staff||staff.role!==role)throw Error('同じ職種の中で並べ替えてください。');
    const members=state.staff.filter(s=>s.active&&s.role===role),old=members.findIndex(s=>s.id===id);
    if(old<index)index--;
    state.staff=state.staff.filter(s=>s.id!==id);
    const remaining=state.staff.filter(s=>s.active&&s.role===role),target=remaining[Math.max(0,Math.min(index,remaining.length))];
    const at=target?state.staff.indexOf(target):(remaining.length?state.staff.indexOf(remaining.at(-1))+1:state.staff.length);
    state.staff.splice(at,0,staff);
  }
  function resolve(state,issueId,targetId,adoptName=false) {
    const q=state.issues.find(x=>x.id===issueId);if(!q || q.resolved)throw Error('確認項目が見つかりません。');
    if(q.kind==='link') {
      const p=state.residents.find(p=>p.id===targetId && p.active);if(!p)throw Error('対応する入所者を選んでください。');
      if(q.group && !q.group.startsWith('laundry-') && state.orders[q.group].includes(p.id))throw Error('選んだ人はすでに同じ枠にいます。対応する別のカードを確認してください。');
      Object.assign(p,copy(q.payload));
      if(q.payload.nameColor)p.gender={red:'female',blue:'male',black:''}[q.payload.nameColor]||'';
      if(adoptName)p.name=q.sourceName;
      if(q.group?.startsWith('laundry-'))state.orders[q.group]=state.orders[q.group].filter(id=>id!==p.id);
      if(q.group)state.orders[q.group]=state.orders[q.group].map(x=>x==='?'+q.id?p.id:x);
      q.targetId=p.id;
    }
    q.resolved=true;
  }
  function problems(state,beds) {
    const out=[];const ids=new Set();const occupied=new Map();
    for(const p of state.residents) {
      if(ids.has(p.id))out.push({type:'error',text:'カードIDが重複しています。'});ids.add(p.id);
      if(!p.active)continue;
      if(!p.name?.trim())out.push({type:'error',text:'氏名が空のカードがあります。',id:p.id});
      if(!beds.includes(p.bed))out.push({type:'error',text:`${p.name}：ベッドを設定してください。`,id:p.id});
      if(occupied.has(p.bed))out.push({type:'error',text:`${p.bed}：ベッドが重複しています。`,id:p.id});occupied.set(p.bed,p.id);
      if(!['談話','A','B','経管','禁食'].includes(p.meal))out.push({type:'error',text:`${p.name}：食事区分を設定してください。`,id:p.id});
    }
    for(const q of state.issues)if(!q.resolved)out.push({type:'import',text:q.text,issueId:q.id});
    for(const family of ['bath','care','laundry']) {
      const seen=new Set();
      for(const [key,g] of Object.entries(GROUPS))if(g.family===family){
        const a=state.orders[key]||[];
        if(a.length>g.cap)out.push({type:'error',text:`${g.label}：用紙の人数上限を超えています。`});
        for(const id of a){
          if(id.startsWith('?')||isBathBlank(id)||isCareBlank(id))continue;
          if(seen.has(id))out.push({type:'error',text:`${family}：カードが重複しています。`,id});seen.add(id);
          const p=state.residents.find(p=>p.id===id);
          if(!p || !eligible(p,family))out.push({type:'error',text:`${g.label}に対象外のカードがあります。`,id});
        }
      }
      for(const p of state.residents)if(eligible(p,family)&&!seen.has(p.id))out.push({type:'missing',family,text:`${p.name}：${{bath:'入浴',care:'排泄',laundry:'洗濯'}[family]}が未配置です。`,id:p.id});
    }
    return out;
  }
  function validateImport(s) {
    if(!s || s.schemaVersion!==1 || !Array.isArray(s.residents) || !Array.isArray(s.staff) || !s.orders || !Array.isArray(s.issues))throw Error('このアプリのカードDBではありません。');
    if(s.residents.length>500 || s.staff.length>500)throw Error('データ件数が上限を超えています。');
    const ids=new Set();
    for(const p of s.residents){
      if(typeof p.id!=='string'||ids.has(p.id)||typeof p.name!=='string'||typeof p.bed!=='string'||!p.bath||!p.assigned)throw Error('カードの形式が正しくありません。');ids.add(p.id);
      if(p.weightReferences!==undefined&&(p.weightReferences===null||Array.isArray(p.weightReferences)||typeof p.weightReferences!=='object'||Object.entries(p.weightReferences).some(([month,value])=>!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)||typeof value!=='number'||!Number.isFinite(value)||value<=0)))throw Error('参照体重の形式が正しくありません。');
      for(const key of ['active','balloon','portable','toilet','bathEligible','careEligible'])if(typeof p[key]!=='boolean')throw Error('カードの設定形式が正しくありません。');
    }
    for(const key of Object.keys(GROUPS))if(!Array.isArray(s.orders[key]) || s.orders[key].length>GROUPS[key].cap || s.orders[key].some(id=>typeof id!=='string'||(!ids.has(id)&&!(key.startsWith('bath-')&&isBathBlank(id))&&!(key.startsWith('care-')&&isCareBlank(id)&&key.endsWith('-'+roomGroup({bed:careBlankRoom(id)+'-1'})))&&!s.issues.some(q=>'?'+q.id===id&&!q.resolved))))throw Error('配置データが正しくありません。');
    return syncRoomGroups(copy(s));
  }
  function removeFromOrders(s,p) {
    for(const [key,g] of Object.entries(GROUPS))if(!eligible(p,g.family))s.orders[key]=s.orders[key].filter(id=>id!==p.id);
  }
  function careCells(p) {
    return {remark:[p.balloon?'バルーン':'',p.portable?'P':''].filter(Boolean).join('・'),
      checks:p.balloon?['なし','','破棄']:p.toilet?['T','T','T']:['','',''],gray:!p.toilet||p.balloon};
  }
  function bathDay(date) {const d=new Date(date+'T12:00:00').getDay();return [1,4].includes(d)?'mon':[2,5].includes(d)?'tue':null;}
  root.WardCore={GROUPS,copy,eligible,move,moveStaff,resolve,problems,validateImport,removeFromOrders,careCells,bathDay,roomGroup,syncRoomGroups,room,isBathBlank,isCareBlank,careBlankRoom};
})(globalThis);
