(async function(){
 'use strict';
 const config=window.WARD_PUBLIC_CONFIG,core=window.WardCore,excel=window.WardExcel;
 const $=id=>document.getElementById(id),keys=Object.keys(config.meta),required={
  care:['排泄ケアチェック','患者一覧','リスト'],
  bath:['月木am','月木pm','火金am','火金pm'],
  laundry:['Ａグループ','Ｂグループ'],map:['２階'],food:['元データ','談話','Ａ','Ｂ']
 };
 let dataDir=null,db=null,templates=null,templateFolderName='',busy=false;
 const setStatus=text=>{$('web-status').textContent=text;};
 const ready=()=>{$('web-launch').disabled=!db||!templates||busy;$('web-data-name').textContent=db?(dataDir?dataDir.name:'選択したJSONファイル'):'未選択';$('web-template-name').textContent=templates?templateFolderName:'未選択';};
 const fail=e=>{setStatus(e.message||String(e));ready();};
 async function readDB(file){const parsed=JSON.parse(await file.text());return core.validateImport(parsed);}
 function toBase64(bytes){let binary='';for(let i=0;i<bytes.length;i+=0x8000)binary+=String.fromCharCode(...bytes.subarray(i,i+0x8000));return btoa(binary);}
 async function readTemplates(files){const result={};for(const key of keys){const name=config.meta[key].filename,file=files.get(name);if(!file)throw Error('Excel原本が足りません：'+name);const bytes=new Uint8Array(await file.arrayBuffer()),book=await excel.read(bytes);for(const sheet of required[key])if(!book.sheets[sheet])throw Error(name+' に必要なシート「'+sheet+'」がありません。');result[key]=toBase64(bytes);}return result;}
 async function folderTemplates(dir){const files=new Map();for(const key of keys){const name=config.meta[key].filename;files.set(name,await(await dir.getFileHandle(name)).getFile());}return readTemplates(files);}
 async function chooseData(){if(!window.showDirectoryPicker)throw Error('このブラウザはフォルダ選択に対応していません。下のJSONファイル選択を使ってください。');const dir=await showDirectoryPicker({mode:'readwrite',id:'ward-cards-web-data'}),file=await(await dir.getFileHandle('カードDB.json')).getFile();const state=await readDB(file);dataDir=dir;db=state;setStatus('カードDBを読み込みました。Excel原本を選んでください。');try{templates=await folderTemplates(dir);templateFolderName=dir.name;setStatus('カードDBとExcel原本5ファイルを確認しました。');}catch{templates=null;}ready();}
 async function chooseTemplates(){if(!window.showDirectoryPicker)throw Error('このブラウザはフォルダ選択に対応していません。下のExcelファイル選択を使ってください。');const dir=await showDirectoryPicker({mode:'read',id:'ward-cards-web-templates'});templates=await folderTemplates(dir);templateFolderName=dir.name;setStatus('Excel原本5ファイルを確認しました。');ready();}
 async function writeFile(dir,name,contents){const handle=await dir.getFileHandle(name,{create:true}),writer=await handle.createWritable();try{await writer.write(contents);await writer.close();}catch(e){try{await writer.abort();}catch{}throw e;}}
 function installAutosave(dir,startingDB){let queue=Promise.resolve(),pending=0,backedUp=false,lastRevision=startingDB.revision;
  window.WARD_WEB_PENDING=false;window.WARD_WEB_SAVE_ERROR='';
  window.WARD_WEB_AUTOSAVE=record=>{const json=JSON.stringify(record,null,2);pending++;window.WARD_WEB_PENDING=true;window.WARD_WEB_STATUS?.();
   const task=queue.catch(()=>{}).then(async()=>{const handle=await dir.getFileHandle('カードDB.json'),oldFile=await handle.getFile(),old=core.validateImport(JSON.parse(await oldFile.text()));if(old.revision!==lastRevision)throw Error('カードDBが別の画面で変更されています。上書きせず、現在のデータをファイル保存してください');
    if(!backedUp){const backup=await dir.getDirectoryHandle('DBバックアップ',{create:true}),stamp=new Date().toISOString().replace(/[:.]/g,'-');await writeFile(backup,'カードDB_'+stamp+'.json',await oldFile.arrayBuffer());backedUp=true;}
    await writeFile(dir,'カードDB.json',json);lastRevision=record.revision;window.WARD_WEB_SAVE_ERROR='';});
   queue=task;return task.catch(e=>{window.WARD_WEB_SAVE_ERROR=e.message;throw e;}).finally(()=>{pending--;window.WARD_WEB_PENDING=pending>0;window.WARD_WEB_STATUS?.();});
  };
  window.addEventListener('beforeunload',e=>{if(pending){e.preventDefault();e.returnValue='';}});
 }
 async function launch(){if(!db||!templates||busy)return;busy=true;ready();try{
   const memory=new Map([['ward-cards-db-v1',JSON.stringify(db)],['ward-cards-output-v1',JSON.stringify(db.outputBaseline||{})]]);
   window.WARD_STORAGE={getItem:key=>memory.get(key)??null,setItem:(key,value)=>memory.set(key,String(value)),removeItem:key=>memory.delete(key)};
   window.WARD_SEED={state:db,beds:config.beds,slots:config.slots,meta:config.meta,templates};
   window.WARD_WEB_MODE=true;window.WARD_WEB_DIRECTORY=dataDir;
   if(dataDir)installAutosave(dataDir,db);
   $('web-start').hidden=true;$('web-app').hidden=false;
   const script=document.createElement('script');script.src='app.js';script.onerror=()=>{setStatus('アプリ本体を読み込めませんでした。');$('web-start').hidden=false;$('web-app').hidden=true;};document.body.appendChild(script);
  }catch(e){busy=false;fail(e);}}
 $('web-data-folder').onclick=async()=>{try{await chooseData();}catch(e){if(e.name!=='AbortError')fail(e);}};
 $('web-template-folder').onclick=async()=>{try{await chooseTemplates();}catch(e){if(e.name!=='AbortError')fail(e);}};
 $('web-data-file').onchange=async e=>{try{const file=e.target.files?.[0];if(!file)return;db=await readDB(file);dataDir=null;setStatus('カードDBを読み込みました。フォルダを選んでいないため、保存はJSONダウンロードになります。');ready();}catch(err){fail(err);}};
 $('web-template-files').onchange=async e=>{try{const files=new Map([...e.target.files].map(f=>[f.name,f]));templates=await readTemplates(files);templateFolderName='選択したExcel 5ファイル';setStatus('Excel原本5ファイルを確認しました。');ready();}catch(err){fail(err);}};
 $('web-launch').onclick=launch;
 ready();
})();
