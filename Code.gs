/** Drape60 backend. Open your Google Sheet > Extensions > Apps Script, paste this file, run setup() once, then deploy as a Web app. */
const H={
 Products:['id','store','n','g','c','p','d','imgs','cols','off','updated'],
 Orders:['id','orderId','store','pid','name','color','size','qty','amt','coupon','custId','custName','phone','area','addr','lat','lng','status','placed','updated'],
 Coupons:['code','type','value','cap','min','from','to','desc','status','needStores','approvedBy'],
 Customers:['id','name','email','phone','dob','taste','area','addr','lat','lng','joined','lastSeen'],
 Stores:['id','name','area','addr','email','phone','status'],
 Activity:['time','who','role','action','detail','mins'],
 Settings:['key','value']};
const KEY={Products:'id',Orders:'id',Coupons:'code',Customers:'id',Stores:'id',Settings:'key'};
const RANK={new:0,accepted:1,packing:2,packed:2,dispatched:3,picked:4,done:5,rejected:9};
/* Order emails go to the Gmail that owns this script. To use a different address: Project Settings > Script properties > add OWNER_EMAIL */
const ownerMail=()=>PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL')||Session.getEffectiveUser().getEmail();
const mapLink=(la,ln,q)=>(la&&ln)?'https://www.google.com/maps?q='+la+','+ln:'https://www.google.com/maps/search/?api=1&query='+encodeURIComponent(q||'');
const hist=(old,st)=>{const h=Array.isArray(old)?old.slice():[];h.push({s:st,t:new Date().toISOString()});return h};
const now=()=>new Date().toISOString(), today=()=>Utilities.formatDate(new Date(),'Asia/Kolkata','yyyy-MM-dd');

function setup(){
  const ss=SpreadsheetApp.getActive();
  for(const n in H){const s=ss.getSheetByName(n)||ss.insertSheet(n);if(s.getLastRow()==0){s.appendRow(H[n]);s.setFrozenRows(1)}}
  const sp=PropertiesService.getScriptProperties();
  if(!sp.getProperty('ADMIN_KEY'))sp.setProperty('ADMIN_KEY','change-me-123');   // change this in Project Settings > Script properties
  const me=Session.getEffectiveUser().getEmail();
  if(!rows('Stores').length)[['youvthi','Youvthi','Madhapur'],['tarangg','Tarangg','Kukatpally'],['kisah','Kisah','Manikonda'],['issa-studio','Issa Studio','Gachibowli']]
    .forEach((s,i)=>put('Stores',{id:s[0],name:s[1],area:s[2],addr:s[1]+', '+s[2]+', Hyderabad',email:me,phone:'900000000'+(i+1),status:'active'}));
  if(!rows('Settings').length)put('Settings',{key:'banner',value:'Diwali edit · Code DIWALI15 · Free delivery above ₹999'});
  Logger.log('Setup done. Test stores use your email '+me+' with phones 9000000001 to 9000000004. Admin key: '+sp.getProperty('ADMIN_KEY'));
}

/* ---------- sheet helpers ---------- */
const sh=n=>SpreadsheetApp.getActive().getSheetByName(n);
function rows(n){const v=sh(n).getDataRange().getValues();if(v.length<2)return[];const h=v[0];
  return v.slice(1).map(r=>{const o={};h.forEach((k,i)=>{let x=r[i];
    if(x instanceof Date)x=Utilities.formatDate(x,'Asia/Kolkata',"yyyy-MM-dd'T'HH:mm:ss");
    else if(typeof x=='string'&&/^[\[{]/.test(x)){try{x=JSON.parse(x)}catch(e){}}
    o[k]=x});return o})}
function put(n,o,key){key=key===undefined?KEY[n]:key;const s=sh(n);
  const h=s.getRange(1,1,1,Math.max(1,s.getLastColumn())).getValues()[0];
  Object.keys(o).forEach(k=>{if(o[k]!==undefined&&!h.includes(k)){h.push(k);s.getRange(1,h.length).setValue(k)}});
  const cell=x=>x===undefined?'':(x!==null&&typeof x=='object'?JSON.stringify(x):x);
  let i=-1;if(key&&s.getLastRow()>1){const ids=s.getRange(2,h.indexOf(key)+1,s.getLastRow()-1,1).getValues().flat().map(String);i=ids.indexOf(String(o[key]))}
  if(i>=0){const old=s.getRange(i+2,1,1,h.length).getValues()[0];
    s.getRange(i+2,1,1,h.length).setValues([h.map((k,j)=>o[k]===undefined?old[j]:cell(o[k]))])}
  else s.appendRow(h.map(k=>cell(o[k])));}
function remove(n,id){const s=sh(n),h=s.getRange(1,1,1,s.getLastColumn()).getValues()[0];if(s.getLastRow()<2)return;
  const ids=s.getRange(2,h.indexOf(KEY[n])+1,s.getLastRow()-1,1).getValues().flat().map(String),i=ids.indexOf(String(id));if(i>=0)s.deleteRow(i+2)}
function saveImgs(arr){const it=DriveApp.getFoldersByName('Drape60 Images'),f=it.hasNext()?it.next():DriveApp.createFolder('Drape60 Images');
  return arr.map(x=>{if(typeof x!='string'||x.indexOf('data:')!==0)return x;const m=x.match(/^data:(.+?);base64,(.*)$/);if(!m)return x;
    const file=f.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]),m[1],Utilities.getUuid()+'.jpg'));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
    return 'https://drive.google.com/thumbnail?id='+file.getId()+'&sz=w700'})}

/* ---------- login: ONE OTP sent to email and phone ---------- */
function sendSms(phone,msg){/* Plug in MSG91 / Twilio here using UrlFetchApp and return true when sent. */return false}
function otp(d){const role=d.role,em=String(d.email||'').toLowerCase().trim(),ph=String(d.phone||'').trim();let to={e:em,p:ph},id=em||ph;
  if(role=='storereg'){if(!em||!/^\d{10}$/.test(ph))return{err:'Enter a valid email and 10-digit mobile number'};
    if(rows('Stores').find(s=>String(s.email).toLowerCase()==em))return{err:'This email is already registered. Please log in.'};id=em}
  else if(role=='store'){if(!rows('Stores').find(s=>String(s.email).toLowerCase()==em&&String(s.phone)==ph&&s.status!='disabled'))return{err:'No active store matches this email and phone'};id=em}
  else{if(!id)return{err:'Enter your email or mobile'};
    const c=rows('Customers').find(c=>(em&&String(c.email).toLowerCase()==em)||(ph&&String(c.phone)==ph));if(c)to={e:c.email||em,p:c.phone||ph}}
  const code=String(Math.floor(100000+Math.random()*900000));
  CacheService.getScriptCache().put('otp_'+role+'_'+id,code,600);
  const msg='Your Drape60 login code is '+code+'. It is valid for 10 minutes.';let sent=0,sms=0;
  if(to.e){MailApp.sendEmail(to.e,'Your Drape60 login code',msg);sent=1}
  if(to.p&&sendSms(to.p,msg)){sms=1;sent=1}
  return sent?{ok:1,sms:sms}:{err:'Could not send the code. Please use your email.'}}
function verify(d){const role=d.role,id=String(d.id||'').toLowerCase().trim(),c=CacheService.getScriptCache(),k='otp_'+role+'_'+id;
  if(!d.code||c.get(k)!=String(d.code).trim())return{err:'Incorrect or expired OTP'};c.remove(k);
  const o={role,id},token=Utilities.getUuid();let customer=null;
  if(role=='store'){const st=rows('Stores').find(s=>String(s.email).toLowerCase()==id&&String(s.phone)==String(d.phone||'').trim());if(!st)return{err:'Store not found'};o.id=st.id}
  else if(role=='storereg'){const nm=String(d.name||'Store').trim(),sid=nm.toLowerCase().replace(/[^a-z0-9]+/g,'-')+'-'+Math.floor(1000+Math.random()*9000);
    put('Stores',{id:sid,name:nm,owner:String(d.owner||''),email:id,phone:String(d.phone||'').trim(),status:'active',verifyStatus:'Draft',area:'',addr:'',created:now()});o.role='store';o.id=sid}
  else{const ex=rows('Customers').find(x=>String(x.id).toLowerCase()==id),em=id.indexOf('@')>0;
    put('Customers',{id:id,email:ex?undefined:(em?id:''),phone:ex?undefined:(em?'':id),joined:ex?undefined:now(),lastSeen:now()});
    customer=rows('Customers').find(x=>String(x.id).toLowerCase()==id)}
  c.put('t_'+token,JSON.stringify(o),21600);return{token,customer}}
function orderMail(by){const stores=rows('Stores');
  for(const name in by){const L=by[name],f=L[0],st=stores.find(x=>x.name==name)||{},tot=L.reduce((a,r)=>a+(+r.amt||0),0),pk=f.mode=='pickup';
    const items=L.map(r=>'• '+r.name+' · '+r.color+' · '+r.size+' × '+r.qty+' = ₹'+r.amt).join('\n');
    const drop=pk?'Customer will collect from the store (pickup code '+f.pickupCode+')':f.addr;
    const body='Order #'+f.orderId+' for '+name+'\n\n'+items+'\n\nTotal: ₹'+tot+'\nMode: '+(pk?'Store pickup':'Delivery')+'\nCustomer: '+(f.custName||'')+' · '+(f.phone||f.custId)+'\nDrop: '+drop+(pk?'':'\nDrop map: '+mapLink(f.lat,f.lng,f.addr));
    MailApp.sendEmail(ownerMail(),'New order #'+f.orderId+' · '+name,body+'\n\nPickup from: '+name+', '+(st.addr||'')+(st.phone?' · '+st.phone:''));
    if(st.email)MailApp.sendEmail(st.email,'New order #'+f.orderId+' on Drape60',body)}}
function dispatchMail(o,st){const k=(st.kyc&&st.kyc.s4)||{},pick=st.name+', '+(st.addr||k.storeAddr||''),pm=k.mapLink||mapLink(0,0,pick);
  MailApp.sendEmail(ownerMail(),'DISPATCH · book a bike for order #'+(o.orderId||o.id),
    'Store '+st.name+' has packed order #'+(o.orderId||o.id)+' and pressed Dispatch.\n\nPICKUP\n'+pick+'\nStore phone: '+(st.phone||'')+'\nMap: '+pm+
    '\n\nDROP\n'+(o.custName||'')+' · '+(o.phone||o.custId||'')+'\n'+(o.addr||'')+'\nMap: '+mapLink(o.lat,o.lng,o.addr)+
    '\n\nItem: '+o.name+' · '+o.color+' · '+o.size+' × '+o.qty+'\nOrder value: ₹'+o.amt+
    '\n\nAfter you book the bike, open the admin app > Orders > Manage and enter the rider name, phone and tracking link. The customer will see it live.')}
function auth(t,role){const v=t&&CacheService.getScriptCache().get('t_'+t);if(!v)throw new Error('Please log in again');const o=JSON.parse(v);if(o.role!=role)throw new Error('Not allowed');return o}

/* ---------- stock alert email ---------- */
function lowAlert(st,ps){const l=[],h=[];
  ps.forEach(p=>{const low=[];for(const c in p.cols)for(const z in p.cols[c])if(+p.cols[c][z]<5)low.push(c+' · size '+z+': '+p.cols[c][z]+' left');
    if(low.length){l.push(p.n+' - '+low.join('; '));const im=(p.imgs&&p.imgs[0])||'';
      h.push('<div style="display:flex;gap:12px;margin:12px 0">'+(im?'<img src="'+im+'" width="80" style="border-radius:8px">':'')+'<div><b>'+p.n+'</b><br>'+low.join('<br>')+'</div></div>')}});
  if(!l.length)return;const key='low_'+st.id+'_'+Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,l.join('|'))).slice(0,16),c=CacheService.getScriptCache();
  if(c.get(key))return;c.put(key,'1',21600);
  MailApp.sendEmail({to:st.email,cc:ownerMail(),subject:'Low stock alert · '+st.name,body:'Items below 5 units:\n\n'+l.join('\n'),htmlBody:'<p>These items are below 5 units:</p>'+h.join('')})}
function saveKyc(o,sid){const it=DriveApp.getFoldersByName('Drape60 Verification'),f=it.hasNext()?it.next():DriveApp.createFolder('Drape60 Verification');
  const w=v=>{if(typeof v=='string'&&v.indexOf('data:')===0){const m=v.match(/^data:(.+?);base64,(.*)$/);if(!m)return v;
      const file=f.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]),m[1],sid+'-'+Utilities.getUuid().slice(0,8)+'.'+(m[1].split('/')[1]||'bin')));
      return 'https://drive.google.com/file/d/'+file.getId()+'/view'}
    if(Array.isArray(v))return v.map(w);if(v&&typeof v=='object'){const r={};for(const k in v)r[k]=w(v[k]);return r}return v};
  return w(o)}

/* ---------- router ---------- */
function doGet(){return out({ok:1,msg:'Drape60 API is running'})}
const out=r=>ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
function doPost(e){const lock=LockService.getScriptLock();let r;
  try{lock.waitLock(15000);const q=JSON.parse(e.postData.contents);r=route(q.a,q.d||{},q.t,q.k)}catch(x){r={err:String(x.message||x)}}finally{try{lock.releaseLock()}catch(x){}}
  return out(r)}
function route(a,d,t,k){
  if(a=='boot'){const stores=rows('Stores').filter(s=>s.status!='disabled'&&(s.verifyStatus||'Verified')=='Verified'),names=stores.map(s=>s.name),td=today();
    return{products:rows('Products').filter(p=>names.includes(p.store)),stores:stores.map(s=>({name:s.name,area:s.area})),
      coupons:rows('Coupons').filter(c=>c.status=='approved'&&(!c.needStores||(c.approvedBy||[]).length)&&(!c.from||String(c.from).slice(0,10)<=td)&&(!c.to||String(c.to).slice(0,10)>=td)),
      offer:(rows('Settings').find(s=>s.key=='banner')||{}).value||''}}
  if(a=='otp')return otp(d);if(a=='verify')return verify(d);
  if(a=='activity'){let who=d.who;try{const v=CacheService.getScriptCache().get('t_'+t);if(v)who=who||JSON.parse(v).id}catch(x){}
    put('Activity',{time:now(),who:who,role:d.role,action:d.action,detail:d.detail||'',mins:d.mins||''},null);return{ok:1}}
  if(a=='cust.save'){const s=auth(t,'customer');put('Customers',{id:s.id,name:d.name,dob:d.dob,taste:d.taste,area:d.area,addr:d.addr,lat:d.lat,lng:d.lng,lastSeen:now()});return{ok:1}}
  if(a=='order.place'){const s=auth(t,'customer'),by={};
    d.rows.forEach(r=>{r.custId=s.id;r.status='new';r.updated=now();r.history=[{s:'new',t:now()}];put('Orders',r);(by[r.store]=by[r.store]||[]).push(r)});
    try{orderMail(by)}catch(x){}return{ok:1}}
  if(a=='order.mine'){const s=auth(t,'customer');return{orders:rows('Orders').filter(o=>String(o.custId)==s.id)}}
  if(a=='store.boot'){const s=auth(t,'store'),st=rows('Stores').find(x=>x.id==s.id);
    return{store:st,products:rows('Products').filter(p=>p.store==st.name),orders:rows('Orders').filter(o=>o.store==st.name).reverse().slice(0,2000),
      coupons:rows('Coupons').filter(c=>c.status=='approved'&&c.needStores)}}
  if(a=='store.kyc'){const s=auth(t,'store'),st=rows('Stores').find(x=>x.id==s.id),vs=st.verifyStatus||'Verified';
    if(!['Draft','Rejected','More Documents Required'].includes(vs))return{err:'Your application is '+vs+' and cannot be edited now'};
    const k=saveKyc(d.kyc||{},st.id),row={id:st.id,kyc:k};if(d.submit){row.verifyStatus='Pending';row.submitted=now();row.adminNote=''}
    put('Stores',row);
    if(d.submit)try{MailApp.sendEmail(st.email,'Drape60: verification received','We received your store verification. Our team will review it and email you. This usually takes 1 to 3 working days.')}catch(x){}
    return{kyc:k,verifyStatus:d.submit?'Pending':vs}}
  if(a=='store.push'){const s=auth(t,'store'),st=rows('Stores').find(x=>x.id==s.id),keep=[];
    put('Stores',{id:st.id,area:d.store.area,addr:d.store.addr,email:d.store.email,phone:d.store.phone});
    (d.products||[]).forEach(p=>{p.store=st.name;p.imgs=saveImgs(p.imgs||[]);p.updated=now();put('Products',p);keep.push(String(p.id))});
    rows('Products').filter(p=>p.store==st.name&&!keep.includes(String(p.id))).forEach(p=>remove('Products',p.id));
    const ex={};rows('Orders').filter(x=>x.store==st.name).forEach(x=>ex[String(x.id)]=x);
    (d.orders||[]).forEach(o=>{o.store=st.name;const old=ex[String(o.id)];
      if(old){const a1=old.status,b1=o.status,ok=b1&&!['done','rejected'].includes(a1)&&((b1=='rejected'&&a1=='new')||(b1!='rejected'&&(RANK[b1]||0)>(RANK[a1]||0)));
        if(!ok)o.status=a1;else if(b1!=a1){o.history=hist(old.history,b1);if(b1=='done')o.deliveredAt=now();
          if(b1=='dispatched'&&old.mode!='pickup')try{dispatchMail(Object.assign({},old,o),st)}catch(x){}}}
      else o.history=hist([],o.status||'new');
      o.updated=now();put('Orders',o)});
    lowAlert(st,d.products||[]);return{products:d.products}}
  if(a=='store.approve'){const s=auth(t,'store'),st=rows('Stores').find(x=>x.id==s.id),c=rows('Coupons').find(x=>x.code==d.code);if(!c)return{err:'Coupon not found'};
    let ap=(Array.isArray(c.approvedBy)?c.approvedBy:[]).filter(n=>n!=st.name);if(d.on)ap.push(st.name);put('Coupons',{code:c.code,approvedBy:ap});return{ok:1}}
  if(String(a).indexOf('admin.')==0)return admin(a,d,k);
  return{err:'Unknown action'}}
function admin(a,d,k){
  if(!k||k!=PropertiesService.getScriptProperties().getProperty('ADMIN_KEY'))return{err:'Wrong admin key'};
  if(a=='admin.boot')return{stores:rows('Stores'),orders:rows('Orders').reverse(),coupons:rows('Coupons'),customers:rows('Customers'),
    activity:rows('Activity').slice(-2000),settings:rows('Settings'),products:rows('Products').map(p=>({id:p.id,store:p.store,n:p.n,p:p.p,cols:p.cols}))};
  if(a=='admin.save'){if(!['Stores','Coupons','Settings','Orders'].includes(d.sheet))return{err:'Not allowed'};
    if(d.sheet=='Orders'&&d.row.status){const old=rows('Orders').find(x=>String(x.id)==String(d.row.id));
      if(old&&old.status!=d.row.status){d.row.history=hist(old.history,d.row.status);if(d.row.status=='done')d.row.deliveredAt=now()}d.row.updated=now()}
    put(d.sheet,d.row);
    if(d.sheet=='Stores'&&d.row.verifyStatus){const st=rows('Stores').find(x=>x.id==d.row.id);
      if(st&&st.email)try{MailApp.sendEmail(st.email,'Drape60 verification: '+d.row.verifyStatus,'Your store verification status is now: '+d.row.verifyStatus+(d.row.adminNote?'\n\nNote from our team: '+d.row.adminNote:'')+(d.row.verifyStatus=='Verified'?'\n\nYour store is live and can now receive orders.':''))}catch(x){}}
    return{ok:1}}
  if(a=='admin.del'){if(!['Stores','Coupons'].includes(d.sheet))return{err:'Not allowed'};remove(d.sheet,d.id);return{ok:1}}
  return{err:'Unknown admin action'}}
