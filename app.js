const KEY='fwts_v1';
const $=s=>document.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const RADIUS=300;
let db={rows:[],mgr:{id:'manager',pw:'admin123'},pw:{}};
let S={user:null,role:'manager',tab:'dash',beat:'',demoGps:false,err:''};
try{const r=localStorage.getItem(KEY);if(r)db=JSON.parse(r);db.pw=db.pw||{};db.mgr=db.mgr||{id:'manager',pw:'admin123'}}catch(e){}
const save=()=>{try{localStorage.setItem(KEY,JSON.stringify(db))}catch(e){}};
const ymd=d=>{d=d||new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')};
const TODAY=()=>ymd();
const ciOn=(r,d)=>(r.checkins||[]).find(c=>c.date===d);
const ciToday=r=>ciOn(r,TODAY());
const ciText=c=>c?new Date(c.time).toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'})+(c.d?' ('+c.d+' m)':''):'';
const lastDays=n=>[...Array(n)].map((_,i)=>{const t=new Date();t.setDate(t.getDate()-(n-1-i));return t});
function ciSeed(skipToday){const o=[];for(let i=skipToday?1:0;i<5;i++){const t=new Date();t.setDate(t.getDate()-i);if(i>0)t.setHours(9,10+i,0,0);o.push({date:ymd(t),time:t.toISOString(),d:30+i*7})}return o}
db.rows.forEach(r=>{if(!r.checkins)r.checkins=r.checkin?[{date:ymd(new Date(r.checkin.time)),time:r.checkin.time,d:r.checkin.d,ref:r.checkin.ref}]:[];delete r.checkin});
function toast(m){const t=document.createElement('div');t.className='toast';t.textContent=m;document.body.appendChild(t);setTimeout(()=>t.remove(),3000)}
function dist(a,b,c,d){const R=6371000,t=x=>x*Math.PI/180,dl=t(c-a),dn=t(d-b);const h=Math.sin(dl/2)**2+Math.cos(t(a))*Math.cos(t(c))*Math.sin(dn/2)**2;return 2*R*Math.asin(Math.sqrt(h))}
const noGps=r=>r.lat==null||r.lng==null||r.lat===''||r.lng===''||isNaN(r.lat)||isNaN(r.lng);
const status=r=>r.progress>=100?'COMPLETED':(r.progress>0||r.checkins.length)?'In Progress':'Pending';
const uDet=u=>{const l=[];if(u.act)l.push(`<div>${esc(u.act)}</div>`);if(u.stock||u.issues)l.push(`<div class="sm mut">${u.stock?'Stock: '+esc(u.stock):''}${u.issues?(u.stock?' · ':'')+'Issues: '+esc(u.issues):''}</div>`);const sc=[];if(u.scan!==undefined&&u.scan!=='')sc.push('Count: '+esc(u.scan));if(u.hgp)sc.push('HGP: '+esc(u.hgp));if(u.hgo)sc.push('HGO: '+esc(u.hgo));if(sc.length)l.push(`<div class="sm"><b>No. of scan</b> · ${sc.join(' · ')}</div>`);if(u.rem)l.push(`<div class="sm mut">Remarks: ${esc(u.rem)}</div>`);return l.join('')};
const att=r=>ciToday(r)?'Checked in today':'Not checked in today';
const pc=r=>{const s=status(r);return`<span class="pill ${s==='COMPLETED'?'ok':s==='Pending'?'wa':'ac'}">${s}</span>`};
const bar=p=>`<div class="row" style="flex-wrap:nowrap"><div class="bar" style="flex:1"><i class="${p>=100?'ok':''}" style="width:${p}%"></i></div><b class="sm">${p}%</b></div>`;
const execs=()=>{const m={};db.rows.forEach(r=>{(m[r.execId]=m[r.execId]||{id:r.execId,name:r.execName,mobile:r.execMobile,rows:[]}).rows.push(r)});return Object.values(m).sort((a,b)=>a.name.localeCompare(b.name))};
const beats=()=>[...new Set(db.rows.map(r=>r.district))].sort();

function demo(){db.sheets=[];db.requests=[];
  const c=[['Arun','E101','9840000101','Workshop A','12 Beach Rd, Cuddalore',11.7480,79.7636],['Bala','E102','9840000102','Workshop B','45 Main St, Cuddalore',11.7502,79.7701],['Karthik','E103','9840000103','Workshop C','8 Bazaar St, Cuddalore',11.7421,79.7589],['Ravi','E104','9840000104','Workshop D','3 Temple Rd, Cuddalore',11.7555,79.7742]];
  const prog=[50,75,0,100];
  db.rows=c.map((x,i)=>({id:'r'+i,sno:String(i+1),son:'SON-100'+(i+1),district:'Cuddalore',execName:x[0],execId:x[1],execMobile:x[2],clientMobile:'98400002'+i+'0',workshop:x[3],address:x[4],lat:x[5],lng:x[6],checkins:ciSeed(i===2),progress:prog[i],updates:prog[i]?[{t:new Date().toISOString(),p:prog[i],act:'Initial stock check',stock:'In stock',issues:'',rem:''}]:[]}));
  db.rows.push({id:'r4',sno:'5',son:'SON-1005',district:'Villupuram',execName:'Dinesh',execId:'E105',execMobile:'9840000105',clientMobile:'9840000250',workshop:'Workshop E',address:'22 Gandhi Rd, Villupuram',lat:11.9401,lng:79.4861,checkins:[],progress:0,updates:[]});
  save();
}
function norm(h){return String(h).toLowerCase().replace(/[^a-z0-9]/g,'')}
function parseSheet(file,add){
  const rd=new FileReader();
  rd.onload=e=>{try{
    const wb=XLSX.read(e.target.result,{type:'array'});const js=wb.SheetNames.flatMap(n=>XLSX.utils.sheet_to_json(wb.Sheets[n],{defval:''}));const base=add?db.rows.length:0;
    const map={sno:['sno','slno','srno','serialno','serialnumber','sl','sr','no','number'],son:['sonnumber','son'],district:['districtcode','district','beat','beatdistrict'],execName:['executivename','employeename'],execId:['executiveid','employeeid'],execMobile:['executivemobilenumber','executivemobile','executivemobileno','executivephone','executivecontact','employeemobile','employeemobileno','employeephone'],clientMobile:['clientmobilenumber','clientmobile','clientmobileno','clientphone','clientcontact','clientnumber','customermobile','customermobileno'],workshop:['workshopname','workshop'],address:['address','addresses','workshopaddress','clientaddress','location'],lat:['latitude','lat'],lng:['longitude','lng','lon']};
    const rows=[];
    js.forEach((o,i)=>{const n={};Object.keys(o).forEach(k=>n[norm(k)]=o[k]);const g=f=>{for(const k of map[f])if(n[k]!==undefined&&n[k]!=='')return n[k];if(f!=='sno')for(const key in n)if(n[key]!==''&&map[f].some(a=>a.length>5&&key.includes(a)))return n[key];return''};
      const r={id:'u'+Date.now()+'_'+base+'_'+i,sno:String(g('sno')).trim()||String(base+i+1),son:g('son'),district:String(g('district')).trim(),execName:String(g('execName')).trim(),execId:String(g('execId')).trim(),execMobile:String(g('execMobile')).replace(/\D/g,''),clientMobile:String(g('clientMobile')).trim(),workshop:g('workshop'),address:g('address'),lat:parseFloat(g('lat')),lng:parseFloat(g('lng')),checkins:[],progress:0,updates:[]};
      if(r.execName&&r.workshop){if(!r.execId)r.execId=r.execMobile||r.execName;rows.push(r)}});
    if(!rows.length)throw new Error('No valid rows (need Executive Name and Workshop Name columns).');
    const info={name:file.name,t:new Date().toISOString()};
    if(add){const key=r=>[r.execId,r.workshop,r.district,r.address].map(x=>String(x).trim().toLowerCase()).join('|');const have=new Set(db.rows.map(key)),fresh=rows.filter(r=>{const k=key(r);if(have.has(k))return false;have.add(k);return true}),dup=rows.length-fresh.length;
      db.rows=db.rows.concat(fresh);db.sheets=(db.sheets||[]).concat({...info,n:fresh.length});save();toast(file.name+': '+fresh.length+' row'+(fresh.length==1?'':'s')+' added'+(dup?', '+dup+' duplicate'+(dup>1?'s':'')+' skipped':''));render();return}
    db.rows=rows;db.sheets=[{...info,n:rows.length}];save();S.tab='exec';toast(rows.length+' rows imported');render();
  }catch(err){toast('Upload failed: '+err.message)}};
  rd.readAsArrayBuffer(file);
}
function templateCSV(){
  const h='S.No,SON Number,District Code,Executive Name,Executive ID,Executive Mobile Number,Client Mobile Number,Workshop Name,Address,Latitude,Longitude\n1,SON-2001,Cuddalore,Arun,E101,9840000101,9840000201,Workshop A,"12 Beach Rd, Cuddalore",11.7480,79.7636\n';
  const a=document.createElement('a');a.href='data:text/csv;charset=utf-8,'+encodeURIComponent(h);a.download='client-sheet-template.csv';a.click();
}

/* auth */
function login(e){e.preventDefault();S.err='';
  if(S.role==='manager'){if($('#u').value.trim()===db.mgr.id&&$('#p').value===db.mgr.pw){S.user={role:'manager'};S.tab='dash'}else S.err='Wrong manager ID or password.'}
  else{const id=$('#u').value.trim().toLowerCase(),pw=$('#p').value;const x=execs().find(z=>z.id.toLowerCase()===id&&pw===(db.pw[z.id]||z.mobile));if(x)S.user={role:'emp',id:x.id};else S.err='Wrong executive ID or password.'}
  render();return false}
const logout=()=>{S.user=null;render()};
function loginView(){return`<div class="login"><section class="hero"><img class="herologo" src="${LOGO}" alt="Employee Field-Work Tracking System logo"><h1>Employee Field-Work Tracking System</h1><p>Upload the client's Excel sheet, assign workshops to executives, verify daily attendance by GPS, every day and watch progress reach 100%.</p><div class="flow">Manager Login → Excel Upload → Beat/District → Assignment → Employee Login → Daily 300 m GPS Check-in → Work Update → Progress → COMPLETED</div></section>
<section class="lb"><div class="card"><div class="seg"><button aria-pressed="${S.role==='manager'}" onclick="S.role='manager';S.err='';render()">Manager</button><button aria-pressed="${S.role==='emp'}" onclick="S.role='emp';S.err='';render()">Executive</button></div>
<form onsubmit="return login(event)"><div class="f"><label for="u">${S.role==='manager'?'Manager ID':'Executive ID'}</label><input id="u" required autocomplete="username" value="${S.role==='manager'?'manager':''}" placeholder="${S.role==='emp'?'e.g. E101':''}"></div>
<div class="f"><label for="p">${S.role==='manager'?'Password':'Password (default: your mobile number)'}</label><input id="p" type="password" required autocomplete="current-password"></div>
${S.err?`<p class="note bd" role="alert" style="margin-top:12px">${esc(S.err)}</p>`:''}
<button class="btn p" style="width:100%;margin-top:14px">Sign in</button></form>
<p class="sm mut" style="margin-bottom:0">Demo: manager / admin123. Executive: ID <b>E101</b>, password <b>9840000101</b>.</p></div></section></div>`}

function header(){return`<header class="top"><div class="wrap"><span class="brand"><img src="${LOGO_ICON}" alt="" width="40" height="40">Field-Work Tracker</span><span class="pill ac">${S.user.role==='manager'?'Manager':'Executive'}</span><span class="who">${S.user.role==='manager'?bellBtn():''}<button class="btn s" onclick="logout()">Sign out</button></span></div></header>`}

/* manager */
const TABS=[['dash','Dashboard'],['beat','Beat / District'],['att','Attendance'],['exec','Executives'],['upload','Excel Upload'],['rep','Reports'],['set','Settings']];
function rowsTable(rows,showBeat=true){if(!rows.length)return'<div class="card mut">No records yet. Upload the client Excel sheet to begin.</div>';
return`<div class="tw"><table><thead><tr><th>Executive</th>${showBeat?'<th>Beat/District</th>':''}<th>Workshop</th><th>Attendance</th><th>Progress</th><th>Status</th><th></th></tr></thead><tbody>${rows.map(r=>`<tr><td>${esc(r.execName)}<div class="sm mut">${esc(r.execId)}</div></td>${showBeat?`<td>${esc(r.district)}</td>`:''}<td>${esc(r.workshop)}<div class="sm mut">${esc(r.address)}</div></td><td><span class="pill ${ciToday(r)?'ok':''}">${att(r)}</span></td><td style="min-width:120px">${bar(r.progress)}</td><td>${pc(r)}</td><td><button class="btn s" onclick="detail('${r.id}')">Details</button></td></tr>`).join('')}</tbody></table></div>`}
function detail(id){const r=db.rows.find(x=>x.id===id);
$('#layer').innerHTML=`<div class="scrim" onclick="if(event.target===this)closeM()"><div class="modal" role="dialog" aria-modal="true"><h2>${esc(r.workshop)}</h2><dl class="kv" style="margin:12px 0"><dt>S.No</dt><dd>${esc(r.sno)}</dd><dt>SON Number</dt><dd>${esc(r.son)}</dd><dt>Beat/District</dt><dd>${esc(r.district)}</dd><dt>Executive</dt><dd>${esc(r.execName)} (${esc(r.execId)}) · ${esc(r.execMobile)}</dd><dt>Client mobile</dt><dd>${esc(r.clientMobile)}</dd><dt>Address</dt><dd>${esc(r.address)}</dd><dt>Today</dt><dd>${ciToday(r)?'Checked in at '+ciText(ciToday(r)):'Not checked in today'}</dd><dt>Days checked in</dt><dd>${r.checkins.length}</dd><dt>Last 7 days</dt><dd>${dots(r)}</dd><dt>Status</dt><dd>${pc(r)} ${r.progress}%</dd></dl><h3>Update history</h3>${r.updates.length?`<ul class="feed">${r.updates.slice().reverse().map(u=>`<li><b>Submitted</b> · <span class="sm mut">${new Date(u.t).toLocaleString()}</span>${uDet(u)}</li>`).join('')}</ul>`:'<p class="mut">No updates submitted.</p>'}<button class="btn" style="margin-top:12px" onclick="closeM()">Close</button></div></div>`}
const closeM=()=>$('#layer').innerHTML='';
function dashV(){const R=db.rows,c=R.filter(r=>status(r)==='COMPLETED').length,ip=R.filter(r=>status(r)==='In Progress').length,ci=R.filter(ciToday).length;
return`<div class="stack"><div class="stats"><div class="stat"><b>${R.length}</b><span>Workshops</span></div><div class="stat"><b>${ci}</b><span>Checked in today</span></div><div class="stat wa"><b>${ip}</b><span>In progress</span></div><div class="stat bd"><b>${R.length-c-ip}</b><span>Pending</span></div><div class="stat ok"><b>${c}</b><span>Completed</span></div></div><div class="row sp"><h2>All field work</h2><button class="btn p s" onclick="exportUpdates('xlsx')" title="Download all executive updates">⬇ Download updates</button></div>${rowsTable(R)}</div>`}
function beatV(){const B=beats();if(!B.length)return'<div class="card mut">No Beat/District data yet. Upload an Excel sheet first.</div>';if(!S.beat||!B.includes(S.beat))S.beat=B[0];
const R=db.rows.filter(r=>r.district===S.beat),ex=execs().filter(e=>e.rows.some(r=>r.district===S.beat));
return`<div class="stack"><div class="f" style="max-width:420px"><label for="beatSel">Beat / District</label><select id="beatSel" onchange="S.beat=this.value;render()">${B.map(b=>{const t=db.rows.filter(r=>r.district===b);return`<option value="${esc(b)}"${b===S.beat?' selected':''}>${esc(b)} (${t.filter(ciToday).length}/${t.length})</option>`}).join('')}</select></div>
<h2>${esc(S.beat)} (${R.filter(ciToday).length}/${R.length})</h2><p class="mut" style="margin:0">${R.filter(ciToday).length} of ${R.length} checked in today · ${ex.length} executives · ${R.length} workshops · ${R.filter(r=>status(r)==='COMPLETED').length} completed</p>${rowsTable(R,false)}
<h3>Client details</h3><div class="tw"><table><thead><tr><th>S.No</th><th>Workshop</th><th>Address</th><th>Client mobile</th><th>Executive mobile</th></tr></thead><tbody>${R.map(r=>`<tr><td>${esc(r.sno)}</td><td>${esc(r.workshop)}</td><td>${esc(r.address)}</td><td>${esc(r.clientMobile)}</td><td>${esc(r.execMobile)}</td></tr>`).join('')}</tbody></table></div></div>`}
function attV(){const d=S.day||TODAY();S.day=d;const rows=db.rows.map(r=>({r,c:ciOn(r,d)})),n=rows.filter(x=>x.c).length,days=lastDays(14),E=execs();
return`<div class="stack"><div class="row sp"><div class="f" style="width:220px"><label for="dy">Date</label><input id="dy" type="date" max="${TODAY()}" value="${d}" onchange="S.day=this.value||TODAY();render()"></div><div class="row"><span class="pill ok">${n} checked in</span><span class="pill wa">${rows.length-n} not checked in</span></div></div>
${rows.length?`<div class="tw"><table><thead><tr><th>Executive</th><th>Beat/District</th><th>Workshop</th><th>Check-in</th></tr></thead><tbody>${rows.map(({r,c})=>`<tr><td>${esc(r.execName)}<div class="sm mut">${esc(r.execId)}</div></td><td>${esc(r.district)}</td><td>${esc(r.workshop)}</td><td>${c?`<span class="pill ok">${ciText(c)}</span>`:'<span class="pill">Not checked in</span>'}</td></tr>`).join('')}</tbody></table></div>`:'<div class="card mut">No records yet. Upload the client Excel sheet to begin.</div>'}
<h3>Last 14 days by executive</h3><div class="tw"><table><thead><tr><th>Executive</th>${days.map(t=>`<th>${t.getDate()}</th>`).join('')}<th>Days</th></tr></thead><tbody>${E.map(e=>{const hit=days.map(t=>e.rows.some(r=>ciOn(r,ymd(t))));return`<tr><td>${esc(e.name)}</td>${hit.map(h=>`<td>${h?'<span class="pill ok">✓</span>':'<span class="mut">–</span>'}</td>`).join('')}<td><b>${hit.filter(Boolean).length}/14</b></td></tr>`}).join('')||'<tr><td class="mut">No data</td></tr>'}</tbody></table></div></div>`}
function execV(){const E=execs();if(!E.length)return'<div class="card mut">No executives yet.</div>';
return`<div class="stack"><h2>Executives (A–Z)</h2><div class="tw"><table><thead><tr><th>Name</th><th>ID</th><th>Mobile</th><th>Assigned workshops</th><th>Avg progress</th></tr></thead><tbody>${E.map(e=>`<tr><td><b>${esc(e.name)}</b></td><td>${esc(e.id)}</td><td>${esc(e.mobile)}</td><td>${e.rows.map(r=>`<div>${esc(r.workshop)} <span class="sm mut">· ${esc(r.district)}</span></div>`).join('')}</td><td style="min-width:120px">${bar(Math.round(e.rows.reduce((a,r)=>a+r.progress,0)/e.rows.length))}</td></tr>`).join('')}</tbody></table></div></div>`}
function uploadV(){return`<div class="stack"><div class="card"><h3>Upload client Excel sheet</h3><p class="mut" style="margin-top:0">Accepts .xlsx, .xls or .csv up to 5 MB. Expected columns: S.No, SON Number, District Code, Executive Name, Executive ID, Executive Mobile Number, Client Mobile Number, Workshop Name, Address. Optional: Latitude, Longitude (needed for GPS check-in).</p>
<div class="row"><input type="file" id="fl" accept=".xlsx,.xls,.csv" aria-label="Excel file" onchange="pick(this)" style="flex:1;min-width:200px"><button type="button" class="addb" title="Add another Excel sheet" aria-label="Add another Excel sheet" onclick="$('#fl2').click()">+</button><input type="file" id="fl2" hidden multiple accept=".xlsx,.xls,.csv" onchange="pickAdd(this)"></div><p class="sm mut" style="margin:8px 0 0">Choose a file to replace all data, or tap + to add more Excel sheets to the existing data.</p><div class="row" style="margin-top:12px"><button class="btn" onclick="templateCSV()">Download template</button><button class="btn" onclick="if(confirm('Replace all data with demo data?')){demo();render()}">Load demo data</button><button class="btn" onclick="if(confirm('Delete all data?')){db.rows=[];db.sheets=[];db.requests=[];save();render()}">Clear data</button></div></div>${(db.sheets||[]).length?`<div class="card"><h3>Uploaded sheets</h3><ul class="feed">${db.sheets.map(x=>`<li><b>${esc(x.name)}</b> · ${x.n} row${x.n==1?'':'s'} <span class="sm mut">${new Date(x.t).toLocaleString()}</span></li>`).join('')}</ul></div>`:''}<div class="note">Choosing a file replaces current data; the + button adds sheets without replacing it (duplicate rows are skipped). Data is stored in this browser only; a production deployment needs a server and database.</div></div>`}
function pickAdd(i){const fs=[...i.files];i.value='';fs.forEach(f=>{if(!/\.(xlsx|xls|csv)$/i.test(f.name))return toast('Only Excel or CSV files are allowed');if(f.size>5e6)return toast(f.name+' is larger than 5 MB');parseSheet(f,true)})}
function pick(i){const f=i.files[0];if(!f)return;if(!/\.(xlsx|xls|csv)$/i.test(f.name))return toast('Only Excel or CSV files are allowed');if(f.size>5e6)return toast('File is larger than 5 MB');parseSheet(f,false)}
function repV(){const B=beats(),all=db.rows.flatMap(r=>r.updates.map(u=>({...u,r}))).sort((a,b)=>b.t.localeCompare(a.t)).slice(0,15);
return`<div class="stack"><div class="card"><div class="row sp"><div><h3 style="margin:0">Download executive updates</h3><p class="mut sm" style="margin:4px 0 0">Everything executives submitted (HGP, HGO, remarks, check-in, workshop and beat details) in one file.</p></div><div class="row"><button class="btn p" onclick="exportUpdates('xlsx')">⬇ Download Excel</button><button class="btn" onclick="exportUpdates('csv')">CSV</button></div></div></div><h2>Progress by Beat/District</h2><div class="tw"><table><thead><tr><th>Beat/District</th><th>Workshops</th><th>Pending</th><th>In progress</th><th>Completed</th></tr></thead><tbody>${B.map(b=>{const R=db.rows.filter(r=>r.district===b),n=s=>R.filter(r=>status(r)===s).length;return`<tr><td>${esc(b)}</td><td>${R.length}</td><td>${n('Pending')}</td><td>${n('In Progress')}</td><td>${n('COMPLETED')}</td></tr>`}).join('')||'<tr><td colspan="5" class="mut">No data</td></tr>'}</tbody></table></div><div class="card"><h3>Recent activity</h3>${all.length?`<ul class="feed">${all.map(u=>`<li><b>${esc(u.r.execName)}</b> · ${esc(u.r.workshop)} <span class="sm mut">${new Date(u.t).toLocaleString()}</span>${uDet(u)}</li>`).join('')}</ul>`:'<p class="mut">No updates yet.</p>'}</div></div>`}

/* download executive updates */
const csvSafe=v=>{let t=String(v??'');if(/^[=+\-@\t\r]/.test(t))t="'"+t;return t};
function updateSheets(){
  const hasOld=k=>db.rows.some(r=>r.updates.some(u=>u[k]!==undefined&&u[k]!==''));
  const extra=[['act','Activity'],['stock','Stock'],['issues','Issues'],['scan','Scan Count']].filter(x=>hasOld(x[0]));
  const H=['S.No','SON Number','Beat/District','Executive Name','Executive ID','Executive Mobile','Workshop','Address','Client Mobile','Check-in Time','Check-in Distance (m)','Update Date','Update Time','HGP','HGO',...extra.map(x=>x[1]),'Remarks','Progress %','Status'];
  const list=[];db.rows.forEach(r=>r.updates.forEach(u=>list.push({r,u})));
  list.sort((a,b)=>String(b.u.t).localeCompare(String(a.u.t)));
  const A=[H].concat(list.map(({r,u})=>[r.sno,r.son,r.district,r.execName,r.execId,r.execMobile,r.workshop,r.address,r.clientMobile,
    (c=>c?new Date(c.time).toLocaleString():'No check-in that day')(ciOn(r,ymd(new Date(u.t)))),(c=>c?c.d:'')(ciOn(r,ymd(new Date(u.t)))),
    new Date(u.t).toLocaleDateString(),new Date(u.t).toLocaleTimeString(),u.hgp||'',u.hgo||'',
    ...extra.map(x=>u[x[0]]??''),u.rem||'',u.p??r.progress,status(r)]));
  const H2=['S.No','SON Number','Beat/District','Executive Name','Executive ID','Executive Mobile','Workshop','Address','Client Mobile','Attendance','Check-in Time','Progress %','Status','No. of Updates','Latest HGP','Latest HGO','Latest Remarks'];
  const B=[H2].concat(db.rows.map(r=>{const u=r.updates[r.updates.length-1]||{};return[r.sno,r.son,r.district,r.execName,r.execId,r.execMobile,r.workshop,r.address,r.clientMobile,att(r),(c=>c?new Date(c.time).toLocaleString():'')(ciToday(r)),r.progress,status(r),r.updates.length,u.hgp||'',u.hgo||'',u.rem||'']}));
  const H3=['Date','Check-in Time','Beat/District','Executive Name','Executive ID','Workshop','Address','Distance (m)','Note'],K=[];
  db.rows.forEach(r=>r.checkins.forEach(c=>K.push([c.date,new Date(c.time).toLocaleTimeString(),r.district,r.execName,r.execId,r.workshop,r.address,c.d,c.ref?'Set workshop GPS reference':''])));
  K.sort((a,b)=>b[0].localeCompare(a[0])||a[3].localeCompare(b[3]));
  return{A,B,C:[H3].concat(K)};
}
function saveFile(href,name){const a=document.createElement('a');a.href=href;a.download=name;document.body.appendChild(a);a.click();a.remove()}
function exportUpdates(fmt){
  const{A,B,C}=updateSheets(),n=A.length-1,stamp=new Date().toISOString().slice(0,10);
  if(!n&&C.length<2)return toast('No updates or check-ins yet');
  if(fmt==='csv'||typeof XLSX==='undefined'){
    const csv='\ufeff'+A.map(row=>row.map(c=>'"'+csvSafe(c).replace(/"/g,'""')+'"').join(',')).join('\r\n');
    saveFile('data:text/csv;charset=utf-8,'+encodeURIComponent(csv),'executive-updates-'+stamp+'.csv');
    return toast(fmt==='csv'?'Downloaded '+n+' update'+(n>1?'s':'')+' (CSV)':'Excel library not loaded, so a CSV was downloaded instead')}
  const wb=XLSX.utils.book_new(),add=(rows,name)=>{const ws=XLSX.utils.aoa_to_sheet(rows);
    ws['!cols']=rows[0].map((h,i)=>({wch:Math.min(45,rows.reduce((m,r)=>Math.max(m,String(r[i]??'').length),String(h).length)+2)}));
    ws['!freeze']={xSplit:0,ySplit:1};XLSX.utils.book_append_sheet(wb,ws,name)};
  add(A,'Executive Updates');add(B,'Workshop Summary');add(C,'Daily Check-ins');
  XLSX.writeFile(wb,'executive-updates-'+stamp+'.xlsx');
  toast('Downloaded '+n+' update'+(n>1?'s':'')+' (Excel)');
}

function setV(){const E=execs();
return`<div class="stack"><h2>Settings</h2><div class="card"><h3>Manager (admin) login</h3><form onsubmit="return saveMgr(event)"><div class="g2"><div class="f"><label for="mi">Manager ID</label><input id="mi" required value="${esc(db.mgr.id)}"></div><div class="f"><label for="mn">New password</label><input id="mn" type="password" minlength="6" autocomplete="new-password" placeholder="Leave blank to keep current"></div></div><div class="f"><label for="mc">Current password (required)</label><input id="mc" type="password" required autocomplete="current-password"></div><button class="btn p" style="margin-top:12px">Save manager login</button></form></div>
<div class="card"><h3>Executive logins</h3><p class="mut sm" style="margin-top:0">Each executive signs in with ID and password. Until you set a password, it is their mobile number.</p>${E.length?`<div class="tw"><table><thead><tr><th>Name</th><th>Executive ID</th><th>New password</th><th></th></tr></thead><tbody>${E.map((e,i)=>`<tr><td>${esc(e.name)}<div class="sm mut">${db.pw[e.id]?'Custom password set':'Password = mobile'}</div></td><td><input id="ei${i}" value="${esc(e.id)}" aria-label="ID for ${esc(e.name)}"></td><td><input id="ep${i}" type="password" minlength="4" placeholder="Leave blank to keep" autocomplete="new-password" aria-label="Password for ${esc(e.name)}"></td><td class="row" style="flex-wrap:nowrap"><button class="btn s p" onclick="saveEx(${i})">Save</button><button class="btn s" onclick="resetEx(${i})">Reset</button></td></tr>`).join('')}</tbody></table></div>`:'<p class="mut">Upload an Excel sheet to create executives.</p>'}</div>
<div class="note">Logins are stored in this browser only. For real use, passwords must be hashed and checked on a server.</div></div>`}
function saveMgr(e){e.preventDefault();if($('#mc').value!==db.mgr.pw){toast('Current password is wrong');return false}
db.mgr.id=$('#mi').value.trim();if($('#mn').value)db.mgr.pw=$('#mn').value;save();toast('Manager login updated');render();return false}
function saveEx(i){const e=execs()[i],nid=$('#ei'+i).value.trim(),np=$('#ep'+i).value;
if(!nid)return toast('ID cannot be empty');
if(nid!==e.id&&execs().some(x=>x.id.toLowerCase()===nid.toLowerCase()))return toast('That ID is already used');
if(np&&np.length<4)return toast('Password needs at least 4 characters');
const old=e.id;if(nid!==old){db.rows.forEach(r=>{if(r.execId===old)r.execId=nid});if(db.pw[old]!==undefined){db.pw[nid]=db.pw[old];delete db.pw[old]}}
if(np)db.pw[nid]=np;save();toast('Saved login for '+e.name);render()}
function resetEx(i){const e=execs()[i];delete db.pw[e.id];save();toast('Password reset to mobile number');render()}
function managerView(){const v={dash:dashV,beat:beatV,att:attV,exec:execV,upload:uploadV,rep:repV,set:setV}[S.tab]();
return header()+`<nav class="tabs" role="tablist">${'<div class="wrap">'}${TABS.map(t=>`<button class="tab" role="tab" aria-selected="${S.tab===t[0]}" onclick="S.tab='${t[0]}';render()">${t[1]}</button>`).join('')}</div></nav><main class="wrap">${v}</main>`}

/* employee */
function streak(me){let n=0;const d=new Date(),any=()=>me.rows.some(r=>ciOn(r,ymd(d)));if(!any())d.setDate(d.getDate()-1);while(any()){n++;d.setDate(d.getDate()-1)}return n}
function dots(r){return`<div class="dots" role="img" aria-label="Check-ins over the last 7 days">${lastDays(7).map((t,i)=>{const c=ciOn(r,ymd(t));return`<span class="dot${c?' on':''}${i===6?' td':''}" title="${t.toDateString()}: ${c?'Checked in':'No check-in'}">${t.toLocaleDateString(undefined,{weekday:'narrow'})}</span>`}).join('')}</div>`}
function empView(){const me=execs().find(e=>e.id===S.user.id);if(!me){S.user=null;return loginView()}
const open=me.rows.filter(r=>r.progress<100),done=open.filter(ciToday).length,all=open.length&&done===open.length;
return header()+`<main class="wrap stack"><div class="row sp"><div><h2>Hello, ${esc(me.name)}</h2><p class="mut" style="margin:4px 0 0">${new Date().toLocaleDateString(undefined,{weekday:'long',day:'numeric',month:'long'})} · ${me.rows.length} assigned workshop${me.rows.length>1?'s':''} · ${esc(me.mobile)}</p></div><button class="btn p" onclick="openNew()" aria-label="Add new workshop details for approval">New +</button></div>
<div class="card row sp"><div><h3 style="margin:0">Today's check-in</h3><p class="mut sm" style="margin:4px 0 0">${open.length?done+' of '+open.length+' workshop'+(open.length>1?'s':'')+' checked in. Check in once every day.':'All your workshops are completed.'}</p></div><div class="row"><span class="pill ${all||!open.length?'ok':'wa'}">${all||!open.length?'Done for today':'Check-in pending'}</span><span class="pill ac">${streak(me)}-day streak</span></div></div>${myReqs(me)}
${me.rows.map(r=>`<div class="card"><div class="row sp"><h3 style="margin:0">${esc(r.workshop)}</h3>${pc(r)}</div><dl class="kv" style="margin:10px 0"><dt>Address</dt><dd>${esc(r.address)}</dd><dt>Client mobile</dt><dd>${esc(r.clientMobile)}</dd><dt>Beat/District</dt><dd>${esc(r.district)}</dd><dt>Today</dt><dd>${ciToday(r)?'Checked in at '+ciText(ciToday(r)):'Not checked in today'}</dd><dt>Last 7 days</dt><dd>${dots(r)}</dd></dl>${bar(r.progress)}
<div style="margin-top:12px">${r.progress>=100?'<div class="note ok">Work completed. Well done.</div>':!ciToday(r)?`<button class="btn p" onclick="checkin('${r.id}')">Check in for today</button> <span class="sm mut">${noGps(r)?'First check-in sets this workshop\'s GPS location':'Must be within '+RADIUS+' m'}</span>`:updForm(r)}${locBtn(r)}</div></div>`).join('')}</main>`}
function updForm(r){const o=[25,50,75,100].filter(p=>p>r.progress);return`<form onsubmit="return upd(event,'${r.id}')"><fieldset class="sub"><legend>No. of scan</legend><div class="g2"><div class="f"><label for="pg${r.id}">HGP</label><input id="pg${r.id}"></div><div class="f"><label for="po${r.id}">HGO</label><input id="po${r.id}"></div></div></fieldset><div class="g2" style="margin-top:12px"><div class="f"><label for="pp${r.id}">Progress after today's work</label><select id="pp${r.id}">${o.map(p=>`<option value="${p}">${p===100?'100% (completed)':p+'%'}</option>`).join('')}</select></div><div class="f"><label for="pr${r.id}">Remarks</label><input id="pr${r.id}"></div></div><button class="btn p" style="margin-top:12px">Submit today's update</button></form>`}
function checkin(id){const r=db.rows.find(x=>x.id===id);if(ciToday(r))return toast('Already checked in today');
  const noRef=noGps(r),add=(d,ref)=>{r.checkins.push({date:TODAY(),time:new Date().toISOString(),d,ref:!!ref});save()};
  const ok=(la,ln)=>{if(noRef){r.lat=la;r.lng=ln;add(0,true);toast('Checked in for today. Your location is now saved as this workshop\'s reference GPS.');render();return}
    const d=Math.round(dist(la,ln,r.lat,r.lng));if(d<=RADIUS){add(d);toast('Checked in for today ('+d+' m away)');render()}else toast('You are '+d+' m away. Move within '+RADIUS+' m of the workshop to check in.')};
  if(!navigator.geolocation)return toast('Geolocation is not supported on this device');
  toast('Getting your location…');navigator.geolocation.getCurrentPosition(p=>ok(p.coords.latitude,p.coords.longitude),e=>toast('Could not get location: '+e.message),{enableHighAccuracy:true,timeout:15000})}
function upd(e,id){e.preventDefault();const r=db.rows.find(x=>x.id===id);if(!ciToday(r))return toast('Check in first'),false;const p=+$('#pp'+id).value;
  r.updates.push({t:new Date().toISOString(),p,hgp:$('#pg'+id).value.trim(),hgo:$('#po'+id).value.trim(),rem:$('#pr'+id).value.trim()});r.progress=p;save();toast(p>=100?'Marked COMPLETED':'Update saved ('+p+'%)');render();return false}

/* new workshop requests + approvals */
const reqs=()=>db.requests||(db.requests=[]);
const locs=()=>db.locReqs||(db.locReqs=[]);
const pendingN=()=>reqs().filter(q=>q.status==='pending').length+locs().filter(q=>q.status==='pending').length;
function bellBtn(){const n=pendingN();return`<button class="btn s bell" onclick="openBell()" title="Approvals" aria-label="Approval requests, ${n} pending"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M18 8a6 6 0 0 0-12 0c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.7 21a2 2 0 0 1-3.4 0"/></svg>${n?`<span class="badge">${n}</span>`:''}</button>`}
function myReqs(me){const L=reqs().filter(q=>q.execId===me.id).slice(-5).reverse();if(!L.length)return'';
return`<div class="card"><h3>My new workshop requests</h3><ul class="feed">${L.map(q=>`<li><div class="row sp"><span><b>${esc(q.workshop)}</b> · ${esc(q.district)}</span><span class="pill ${q.status==='approved'?'ok':q.status==='rejected'?'':'wa'}">${q.status==='approved'?'Approved':q.status==='rejected'?'Rejected':'Waiting for approval'}</span></div></li>`).join('')}</ul></div>`}
function openNew(){const me=execs().find(e=>e.id===S.user.id),B=beats(),def=me&&me.rows[0]?me.rows[0].district:B[0];
$('#layer').innerHTML=`<div class="scrim" onclick="if(event.target===this)closeM()"><div class="modal" role="dialog" aria-modal="true" aria-labelledby="nh"><h2 id="nh">New workshop details</h2><p class="mut sm" style="margin:6px 0 12px">The manager will review this before it is added to the beat.</p>
<form onsubmit="return sendNew(event)"><div class="f"><label for="nb">Beat / District</label><select id="nb">${B.map(b=>`<option value="${esc(b)}"${b===def?' selected':''}>${esc(b)}</option>`).join('')}</select></div>
<div class="f"><label for="nw">Workshop name</label><input id="nw" required></div><div class="f"><label for="na">Address</label><textarea id="na" required></textarea></div>
<div class="f"><label for="nm">Client mobile</label><input id="nm" type="tel" inputmode="numeric"></div><div class="f"><label for="nn">Notes</label><input id="nn"></div>
<div class="row" style="margin-top:14px"><button class="btn p">Submit</button><button type="button" class="btn" onclick="closeM()">Cancel</button></div></form></div></div>`;$('#nw').focus()}
function sendNew(e){e.preventDefault();const me=execs().find(x=>x.id===S.user.id);
reqs().push({id:'q'+Date.now(),t:new Date().toISOString(),execId:me.id,execName:me.name,execMobile:me.mobile,district:$('#nb').value,workshop:$('#nw').value.trim(),address:$('#na').value.trim(),clientMobile:$('#nm').value.trim(),notes:$('#nn').value.trim(),status:'pending'});
save();closeM();toast('Sent to the manager for approval');render();return false}
function bellHtml(){const LP=locs().filter(q=>q.status==='pending').reverse(),lcard=q=>`<div class="card" style="margin-top:10px"><div class="row sp"><b>Location reset · ${esc(q.workshop)}</b><span class="sm mut">${new Date(q.t).toLocaleString()}</span></div><p style="margin:6px 0">${esc(q.execName)} (${esc(q.execId)}) says the first check-in was at the wrong location (${esc(q.district)}).</p><dl class="kv"><dt>Reason</dt><dd>${esc(q.reason)}</dd></dl><p class="sm mut" style="margin:8px 0 0">Allow will clear this workshop's saved GPS and the wrong check-in, so the executive can check in again at the correct place.</p><div class="row" style="margin-top:12px"><button class="btn p" onclick="decideLoc('${q.id}',true)">Allow reset</button><button class="btn" onclick="decideLoc('${q.id}',false)">Reject</button></div></div>`;const L=reqs().slice().reverse(),P=L.filter(q=>q.status==='pending'),D=L.filter(q=>q.status!=='pending').slice(0,5);
const card=q=>`<div class="card" style="margin-top:10px"><div class="row sp"><b>${esc(q.workshop)}</b><span class="sm mut">${new Date(q.t).toLocaleString()}</span></div><p style="margin:6px 0">${esc(q.execName)} (${esc(q.execId)}) wants to add a new workshop to <b>${esc(q.district)}</b>.</p><dl class="kv"><dt>Beat/District</dt><dd>${esc(q.district)}</dd><dt>Workshop</dt><dd>${esc(q.workshop)}</dd><dt>Address</dt><dd>${esc(q.address)}</dd><dt>Client mobile</dt><dd>${esc(q.clientMobile)||'—'}</dd><dt>Executive</dt><dd>${esc(q.execName)} · ${esc(q.execMobile)}</dd>${q.notes?`<dt>Notes</dt><dd>${esc(q.notes)}</dd>`:''}</dl>
${q.status==='pending'?`<div class="row" style="margin-top:12px"><button class="btn p" onclick="decide('${q.id}',true)">Approve</button><button class="btn" onclick="decide('${q.id}',false)">Reject</button></div>`:`<div style="margin-top:10px"><span class="pill ${q.status==='approved'?'ok':''}">${q.status==='approved'?'Approved':'Rejected'}</span></div>`}</div>`;
return`<div class="scrim" onclick="if(event.target===this)closeM()"><div class="modal" role="dialog" aria-modal="true" aria-labelledby="bh" style="max-width:560px"><h2 id="bh">Approvals</h2>${LP.map(lcard).join('')}${P.length||LP.length?P.map(card).join(''):'<p class="mut">No requests waiting for approval.</p>'}${D.length?`<h3 style="margin-top:18px">Recent decisions</h3>${D.map(card).join('')}`:''}<button class="btn" style="margin-top:14px" onclick="closeM()">Close</button></div></div>`}
function openBell(){$('#layer').innerHTML=bellHtml()}
function decide(id,ok){const q=reqs().find(x=>x.id===id);if(!q||q.status!=='pending')return;q.status=ok?'approved':'rejected';q.decided=new Date().toISOString();
if(ok)db.rows.push({id:'u'+Date.now()+'_'+q.id,sno:String(db.rows.length+1),son:'',district:q.district,execName:q.execName,execId:q.execId,execMobile:q.execMobile,clientMobile:q.clientMobile,workshop:q.workshop,address:q.address,lat:NaN,lng:NaN,checkins:[],progress:0,updates:[]});
save();toast(ok?q.workshop+' added to '+q.district:'Request rejected');render();openBell()}

function locBtn(r){if(r.progress>=100)return'';const q=locs().filter(x=>x.wid===r.id).pop();
if(q&&q.status==='pending')return'<div class="note" style="margin-top:12px">Location reset requested. Waiting for the manager to allow it.</div>';
const rej=q&&q.status==='rejected'&&ymd(new Date(q.decided))===TODAY()?'<div class="note bd" style="margin-top:12px">The manager rejected your location reset request.</div>':'';
return rej+(ciToday(r)?`<div style="margin-top:12px"><button class="btn s" onclick="openLoc('${r.id}')">Wrong location? Request reset</button></div>`:'')}
function openLoc(id){const r=db.rows.find(x=>x.id===id);
$('#layer').innerHTML=`<div class="scrim" onclick="if(event.target===this)closeM()"><div class="modal" role="dialog" aria-modal="true" aria-labelledby="lh"><h2 id="lh">Request location reset</h2><p class="mut sm" style="margin:6px 0 12px">${esc(r.workshop)} · ${esc(r.address)}. Once the manager allows it, the saved location is cleared and you can check in again at the correct place.</p><form onsubmit="return sendLoc(event,'${id}')"><div class="f"><label for="lr">What went wrong?</label><textarea id="lr" required placeholder="e.g. I checked in from my hotel, not the workshop"></textarea></div><div class="row" style="margin-top:14px"><button class="btn p">Send to manager</button><button type="button" class="btn" onclick="closeM()">Cancel</button></div></form></div></div>`;$('#lr').focus()}
function sendLoc(e,id){e.preventDefault();const r=db.rows.find(x=>x.id===id);if(locs().some(q=>q.wid===id&&q.status==='pending')){closeM();return false}
locs().push({id:'l'+Date.now(),t:new Date().toISOString(),wid:id,execId:r.execId,execName:r.execName,workshop:r.workshop,district:r.district,reason:$('#lr').value.trim(),status:'pending'});
save();closeM();toast('Request sent to the manager');render();return false}
function decideLoc(id,ok){const q=locs().find(x=>x.id===id);if(!q||q.status!=='pending')return;q.status=ok?'approved':'rejected';q.decided=new Date().toISOString();
const r=db.rows.find(x=>x.id===q.wid);
if(ok&&r){const bad=ymd(new Date(q.t));r.checkins=r.checkins.filter(c=>!c.ref&&c.date!==bad);r.lat=NaN;r.lng=NaN}
save();toast(ok?'Location reset for '+q.workshop:'Request rejected');render();openBell()}

function render(){$('#app').innerHTML=!S.user?loginView():S.user.role==='manager'?managerView():empView()}
if(!db.rows.length)demo();
render();
