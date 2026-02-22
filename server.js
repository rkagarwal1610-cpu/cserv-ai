'use strict';
const express  = require('express');
const session  = require('express-session');
const bp       = require('body-parser');
const bcrypt   = require('bcryptjs');
const path     = require('path');
const fs       = require('fs');
const ExcelJS  = require('exceljs');

const app   = express();
const PORT  = process.env.PORT || 3000;
const DB    = path.join(__dirname, 'data.json');
const BKDIR = path.join(__dirname, 'backups');
if (!fs.existsSync(BKDIR)) fs.mkdirSync(BKDIR, { recursive: true });

// ─── DB ───────────────────────────────────────────────────────────────
function load() {
  if (!fs.existsSync(DB)) return initDB();
  try { return JSON.parse(fs.readFileSync(DB, 'utf8')); }
  catch { return initDB(); }
}
function save(d) { fs.writeFileSync(DB, JSON.stringify(d, null, 2)); }

function initDB() {
  const d = {
    users: [
      { id:1, username:'superadmin', password:bcrypt.hashSync('Super@123',10),
        role:'superadmin', name:'Super Administrator', createdAt:new Date().toISOString(),
        permissions:{} },
      { id:2, username:'admin', password:bcrypt.hashSync('Admin@123',10),
        role:'admin', name:'Administrator', createdAt:new Date().toISOString(),
        permissions:{
          roster:{ view:true,generate:true,save:true,export:true,editAgents:true,editRules:true,approve:true },
          shortleave:{ view:true,approve:true,reject:true,cancel:true,dashboard:true }
        }},
      { id:3, username:'operator', password:bcrypt.hashSync('Op@123',10),
        role:'operator', name:'Team Operator', createdAt:new Date().toISOString(),
        permissions:{
          roster:{ view:true },
          shortleave:{ view:true,apply:true }
        }}
    ],
    agents:[
      {emp:'EP0200',name:'Shrikant Nayak',level:0,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0269',name:'Ritu Singh',level:0,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0505',name:'Rohit Kumar Agarwal',level:0,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0563',name:'Himanshi Khowal',level:2,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0564',name:'Chetan Goel',level:1,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0523',name:'Sushant Kumar Suman',level:2,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0560',name:'Mohit Singh',level:3,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0678',name:'Abhay Pratap',level:4,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0726',name:'Swagata Bhoumik',level:6,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0848',name:'Deepak Gupta',level:7,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0442',name:'Shivam Garg',level:1,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0524',name:'Anurag Tiwari',level:4,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0567',name:'Triloki Varshney',level:1,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0557',name:'Sujit Kumar',level:3,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0741',name:'Amarnath Vishwakarma',level:5,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0673',name:'Dhruv Mishra',level:5,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0798',name:'Naveen Kumar S',level:6,dept:'Customer Service',loc:'Gurgaon-US'}
    ],
    rosters:[],
    shortLeaves:[],
    holidays:[],
    notifications:[],  // in-app notifications
    rules:{
      targetWOBase:8, extraWOIfFifthSunday:true,
      maxAgentsOnWOPerDay:2, maxConsecutiveWO:3, maxConsecExtraWOPerAgent:1,
      pairingRule:'senior_junior', minL0Working:true,
      holidayMaxWO:1, holidayMaxAgents:2, allowL0OnHolidayWO:true,
      shortLeaveMonthlyLimit:3
    },
    rosterFormat:{
      showEmp:true,showLevel:true,showRole:true,showDept:false,
      cellWO:'WO',cellROI:'ROI',cellHoliday:'H',cellSunday:'WO',cellSaturday:'WO'
    },
    settings:{ appName:'C-Serv.AI', orgName:'Customer Service Team' },
    moduleAccess:{ roster:true, shortleave:true }
  };
  save(d); return d;
}

// ─── MIDDLEWARE ───────────────────────────────────────────────────────
app.use(bp.json({ limit:'10mb' }));
app.use(bp.urlencoded({ extended:true }));
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({
  secret: process.env.SESSION_SECRET || 'cservai-v3-key-2026',
  resave:false, saveUninitialized:false,
  cookie:{ maxAge:8*60*60*1000, httpOnly:true }
}));

// ─── AUTH ─────────────────────────────────────────────────────────────
const auth  = (q,r,n) => q.session?.user ? n() : r.status(401).json({error:'Login required'});
const isSA  = (q,r,n) => q.session?.user?.role==='superadmin' ? n() : r.status(403).json({error:'Super Admin only'});
const isAdm = (q,r,n) => ['superadmin','admin'].includes(q.session?.user?.role) ? n() : r.status(403).json({error:'Admin only'});
const perm  = (mod,act) => (q,r,n) => {
  const u=q.session?.user;
  if(!u) return r.status(401).json({error:'Login required'});
  if(u.role==='superadmin') return n();
  if(!u.permissions?.[mod]?.[act]) return r.status(403).json({error:'Permission denied'});
  n();
};

// ─── NOTIFICATIONS ─────────────────────────────────────────────────────
function addNotif(d, toId, msg, type='info', refId=null) {
  if(!d.notifications) d.notifications=[];
  d.notifications.unshift({ id:Date.now()+Math.random(), toId, msg, type, refId, read:false, createdAt:new Date().toISOString() });
  // keep last 200
  if(d.notifications.length>200) d.notifications=d.notifications.slice(0,200);
}
app.get('/api/notifications', auth, (req,res)=>{
  const d=load(); const uid=req.session.user.id;
  res.json((d.notifications||[]).filter(n=>n.toId===uid));
});
app.put('/api/notifications/readall', auth, (req,res)=>{
  const d=load(); const uid=req.session.user.id;
  (d.notifications||[]).filter(n=>n.toId===uid).forEach(n=>n.read=true);
  save(d); res.json({ok:true});
});
app.put('/api/notifications/:id/read', auth, (req,res)=>{
  const d=load(); const n=d.notifications?.find(n=>String(n.id)===req.params.id);
  if(n) n.read=true;
  save(d); res.json({ok:true});
});

// ─── AUTH ROUTES ──────────────────────────────────────────────────────
// Public self-registration
app.post('/api/register', (req,res) => {
  const d=load(); const {username,password,name}=req.body;
  if(!username||!password||!name) return res.status(400).json({error:'Name, username and password required'});
  if(password.length<6) return res.status(400).json({error:'Password must be at least 6 characters'});
  if(d.users.find(u=>u.username===username)) return res.status(400).json({error:'Username already taken'});
  const u={id:Date.now(),username,name,password:bcrypt.hashSync(password,10),role:'operator',
    permissions:{ roster:{view:true}, shortleave:{view:true,apply:true} },
    createdAt:new Date().toISOString()};
  d.users.push(u);
  // notify all superadmins/admins
  d.users.filter(u2=>['superadmin','admin'].includes(u2.role))
    .forEach(u2=>addNotif(d,u2.id,`New user registered: ${name} (${username})`, 'user', u.id));
  save(d); res.json({ok:true,id:u.id});
});

app.post('/api/login', (req,res) => {
  const {username,password}=req.body;
  const d=load();
  const u=d.users.find(x=>x.username===username);
  if(!u||!bcrypt.compareSync(password,u.password)) return res.status(401).json({error:'Invalid username or password'});
  req.session.user={id:u.id,username:u.username,role:u.role,name:u.name,permissions:u.permissions||{}};
  res.json({ok:true,user:req.session.user,moduleAccess:d.moduleAccess,settings:d.settings});
});
app.post('/api/logout', (q,r)=>{q.session.destroy();r.json({ok:true});});
app.get('/api/me', auth, (req,res)=>{
  const d=load();
  res.json({user:req.session.user,moduleAccess:d.moduleAccess,settings:d.settings,rules:d.rules,rosterFormat:d.rosterFormat});
});

// ─── AGENTS ───────────────────────────────────────────────────────────
app.get('/api/agents', auth, (_,r)=>r.json(load().agents));
app.put('/api/agents', perm('roster','editAgents'), (req,res)=>{
  const d=load(); d.agents=req.body; save(d); res.json({ok:true});
});

// ─── HOLIDAYS ────────────────────────────────────────────────────────
app.get('/api/holidays', auth, (_,r)=>r.json(load().holidays));
app.put('/api/holidays', isAdm, (req,res)=>{
  const d=load(); d.holidays=req.body; save(d); res.json({ok:true});
});

// ─── RULES ────────────────────────────────────────────────────────────
app.get('/api/rules', auth, (_,r)=>r.json(load().rules));
app.put('/api/rules', isAdm, (req,res)=>{
  const d=load(); d.rules={...d.rules,...req.body}; save(d); res.json({ok:true});
});

// ─── ROSTER FORMAT ────────────────────────────────────────────────────
app.get('/api/rosterformat', auth, (_,r)=>r.json(load().rosterFormat));
app.put('/api/rosterformat', isAdm, (req,res)=>{
  const d=load(); d.rosterFormat={...d.rosterFormat,...req.body}; save(d); res.json({ok:true});
});

// ─── ROSTERS ──────────────────────────────────────────────────────────
app.get('/api/rosters', auth, (req,res)=>{
  const d=load(); const u=req.session.user;
  // Operators only see approved rosters
  const list = ['superadmin','admin'].includes(u.role)
    ? d.rosters
    : d.rosters.filter(r=>r.approved);
  res.json(list.map(r=>({id:r.id,title:r.title,month:r.month,year:r.year,agentCount:r.agentCount,targetWO:r.targetWO,savedAt:r.savedAt,savedBy:r.savedBy,approved:r.approved||false,approvedAt:r.approvedAt||null,approvedBy:r.approvedBy||null})));
});
app.post('/api/rosters', perm('roster','save'), (req,res)=>{
  const d=load();
  const r={...req.body,id:Date.now(),savedAt:new Date().toISOString(),savedBy:req.session.user.name,approved:false};
  d.rosters.unshift(r);
  if(d.rosters.length>100) d.rosters=d.rosters.slice(0,100);
  save(d); res.json({ok:true,id:r.id});
});
app.get('/api/rosters/:id', auth, (req,res)=>{
  const d=load(); const u=req.session.user;
  const r=d.rosters.find(x=>x.id==req.params.id);
  if(!r) return res.status(404).json({error:'Not found'});
  if(!['superadmin','admin'].includes(u.role)&&!r.approved) return res.status(403).json({error:'Roster not yet approved'});
  res.json(r);
});
// Approve roster
app.put('/api/rosters/:id/approve', perm('roster','approve'), (req,res)=>{
  const d=load(); const u=req.session.user;
  const r=d.rosters.find(x=>x.id==req.params.id);
  if(!r) return res.status(404).json({error:'Not found'});
  r.approved=true; r.approvedAt=new Date().toISOString(); r.approvedBy=u.name;
  // notify all operators
  d.users.filter(u2=>u2.role==='operator')
    .forEach(u2=>addNotif(d,u2.id,`Roster "${r.title}" has been approved and is now visible.`,'roster',r.id));
  save(d); res.json({ok:true});
});
app.delete('/api/rosters/:id', isAdm, (req,res)=>{
  const d=load(); d.rosters=d.rosters.filter(x=>x.id!=req.params.id); save(d); res.json({ok:true});
});

// ─── XLSX EXPORT ──────────────────────────────────────────────────────
app.post('/api/rosters/export/xlsx', perm('roster','export'), async (req,res)=>{
  try {
    const {rosterId}=req.body;
    const d=load();
    const rData=d.rosters.find(x=>x.id==rosterId);
    if(!rData) return res.status(404).json({error:'Roster not found'});
    const fmt=d.rosterFormat;
    const MN=['January','February','March','April','May','June','July','August','September','October','November','December'];
    const DS=['Su','Mo','Tu','We','Th','Fr','Sa'];
    const LN=['Team Leader','Sr.L1','Sr.L2','Sr.L3','Jr.L4','Jr.L5','Jr.L6','Trainee'];
    const wb=new ExcelJS.Workbook();
    wb.creator='C-Serv.AI'; wb.created=new Date();
    const ws=wb.addWorksheet(`${MN[rData.month]} ${rData.year}`);
    const fixedCols=[];
    if(fmt.showEmp) fixedCols.push({header:'Emp#',key:'emp',width:10});
    fixedCols.push({header:'Name',key:'name',width:22});
    if(fmt.showLevel) fixedCols.push({header:'Level',key:'level',width:7});
    if(fmt.showRole) fixedCols.push({header:'Role',key:'role',width:14});
    if(fmt.showDept) fixedCols.push({header:'Dept',key:'dept',width:16});
    fixedCols.push({header:'Total WO',key:'two',width:9});
    const dow={};(rData.dowSer||[]).forEach(([k,v])=>dow[+k]=v);
    const dayCols=rData.days.map(d2=>({header:`${DS[dow[d2]??0]}`,key:`d${d2}`,width:5}));
    ws.columns=[...fixedCols,...dayCols];
    const hRow=ws.getRow(1);
    hRow.eachCell(cell=>{
      cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF0D1526'}};
      cell.font={bold:true,color:{argb:'FF00D4FF'},size:10};
      cell.alignment={horizontal:'center',vertical:'middle'};
    });
    hRow.height=20;
    const fixLen=fixedCols.length;
    const dt={};(rData.dtypeSer||[]).forEach(([k,v])=>dt[+k]=v);
    rData.agents.forEach((ag,ri)=>{
      const sc=(rData.schedSer||[]).find(([n])=>n===ag.name)?.[1]||{};
      const rowData={emp:ag.emp,name:ag.name,level:`L${ag.level}`,role:LN[ag.level]||`L${ag.level}`,dept:ag.dept||''};
      let woCount=0;
      rData.days.forEach(d2=>{
        const t=dt[d2]||'WORK',s=sc[d2]||'ROI';
        let val=fmt.cellROI||'ROI';
        if(t==='SUN'||t==='SAT') val=fmt.cellWO||'WO';
        else if(t==='HOL') val=fmt.cellHoliday||'H';
        else if(s==='WO') val=fmt.cellWO||'WO';
        if(val!==(fmt.cellROI||'ROI')) woCount++;
        rowData[`d${d2}`]=val;
      });
      rowData.two=woCount;
      const row=ws.addRow(rowData);
      row.height=18;
      rData.days.forEach((d2,ci)=>{
        const t=dt[d2]||'WORK',s=sc[d2]||'ROI';
        const cell=row.getCell(fixLen+1+ci);
        cell.alignment={horizontal:'center'};
        if(t==='SUN'){cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF2D1F00'}};cell.font={color:{argb:'FFFFB830'},bold:true,size:9}}
        else if(t==='SAT'){cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1A1540'}};cell.font={color:{argb:'FFA78BFA'},bold:true,size:9}}
        else if(t==='HOL'){cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF0D2D1E'}};cell.font={color:{argb:'FF00E5A0'},bold:true,size:9}}
        else if(s==='WO'){cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF2D0A0F'}};cell.font={color:{argb:'FFFF3D5A'},bold:true,size:9}}
        else{cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF060E1A'}};cell.font={color:{argb:'FF4A6080'},size:9}}
      });
      for(let i=1;i<=fixLen;i++){
        const c=row.getCell(i);
        c.fill={type:'pattern',pattern:'solid',fgColor:{argb:ri%2===0?'FF0D1526':'FF080E1A'}};
        c.font={color:{argb:'FFbccde0'},size:10};c.alignment={vertical:'middle'};
      }
    });
    ws.views=[{state:'frozen',xSplit:fixLen,ySplit:1}];
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename="Roster_${MN[rData.month]}_${rData.year}.xlsx"`);
    await wb.xlsx.write(res); res.end();
  } catch(e){console.error(e);res.status(500).json({error:e.message});}
});

// ─── SHORT LEAVE ──────────────────────────────────────────────────────
app.get('/api/shortleaves', auth, (req,res)=>{
  const d=load(); const u=req.session.user;
  const list=['superadmin','admin'].includes(u.role)?d.shortLeaves:d.shortLeaves.filter(sl=>sl.agentId===u.id);
  res.json(list);
});
app.post('/api/shortleaves', perm('shortleave','apply'), (req,res)=>{
  // Synchronous - no await, no email
  const d=load(); const u=req.session.user;
  const {shortLeaveDate,halfDay,reason,requestDate}=req.body;
  if(!shortLeaveDate) return res.status(400).json({error:'Short leave date required'});
  if(!halfDay||!['1st Half','2nd Half'].includes(halfDay)) return res.status(400).json({error:'Please select 1st Half or 2nd Half'});
  const rd=requestDate||new Date().toISOString().slice(0,10);
  const slMonth=shortLeaveDate.slice(0,7);
  const monthCount=d.shortLeaves.filter(sl=>sl.agentId===u.id&&sl.shortLeaveDate.slice(0,7)===slMonth&&sl.status!=='Cancelled').length;
  if(monthCount>=(d.rules.shortLeaveMonthlyLimit||3)) return res.status(400).json({error:`Monthly short leave limit reached for ${slMonth}`});
  const isUnplanned=rd>=shortLeaveDate;
  const sl={
    id:Date.now(),agentId:u.id,agentName:u.name,agentUsername:u.username,
    requestDate:rd,shortLeaveDate,halfDay,reason:reason||'',
    status:'Pending',type:isUnplanned?'Unplanned':'Planned',
    createdAt:new Date().toISOString(),updatedAt:new Date().toISOString(),
    approvedAt:null,rejectedAt:null,cancelledAt:null,
    approvedBy:null,rejectedBy:null,cancelledBy:null,remarks:''
  };
  d.shortLeaves.push(sl);
  // Notify all admins
  d.users.filter(u2=>['superadmin','admin'].includes(u2.role))
    .forEach(u2=>addNotif(d,u2.id,`${u.name} applied Short Leave for ${shortLeaveDate} (${halfDay})`, 'sl_new', sl.id));
  save(d); res.json({ok:true,id:sl.id});
});
app.put('/api/shortleaves/:id/approve', perm('shortleave','approve'), (req,res)=>{
  const d=load(); const u=req.session.user; const {remarks}=req.body;
  const sl=d.shortLeaves.find(x=>x.id==req.params.id);
  if(!sl) return res.status(404).json({error:'Not found'});
  if(sl.status!=='Pending') return res.status(400).json({error:'Only pending requests can be approved'});
  sl.status='Approved'; sl.approvedAt=new Date().toISOString(); sl.approvedBy=u.name; sl.remarks=remarks||''; sl.updatedAt=sl.approvedAt;
  addNotif(d, sl.agentId, `Your Short Leave for ${sl.shortLeaveDate} (${sl.halfDay}) has been Approved by ${u.name}`, 'sl_approved', sl.id);
  save(d); res.json({ok:true});
});
app.put('/api/shortleaves/:id/reject', perm('shortleave','reject'), (req,res)=>{
  const d=load(); const u=req.session.user; const {remarks}=req.body;
  const sl=d.shortLeaves.find(x=>x.id==req.params.id);
  if(!sl) return res.status(404).json({error:'Not found'});
  if(sl.status!=='Pending') return res.status(400).json({error:'Only pending requests can be rejected'});
  sl.status='Rejected'; sl.rejectedAt=new Date().toISOString(); sl.rejectedBy=u.name; sl.remarks=remarks||''; sl.updatedAt=sl.rejectedAt;
  addNotif(d, sl.agentId, `Your Short Leave for ${sl.shortLeaveDate} (${sl.halfDay}) has been Rejected. Reason: ${remarks||'—'}`, 'sl_rejected', sl.id);
  save(d); res.json({ok:true});
});
app.put('/api/shortleaves/:id/cancel', auth, (req,res)=>{
  const d=load(); const u=req.session.user; const {remarks}=req.body;
  const sl=d.shortLeaves.find(x=>x.id==req.params.id);
  if(!sl) return res.status(404).json({error:'Not found'});
  const isAdminRole=['superadmin','admin'].includes(u.role);
  if(!isAdminRole&&sl.agentId!==u.id) return res.status(403).json({error:'Cannot cancel others requests'});
  if(!isAdminRole&&sl.status!=='Pending') return res.status(400).json({error:'You can only cancel pending requests'});
  if(isAdminRole&&!['Pending','Approved'].includes(sl.status)) return res.status(400).json({error:'Can only cancel Pending or Approved requests'});
  sl.status='Cancelled'; sl.cancelledAt=new Date().toISOString(); sl.cancelledBy=u.name; sl.remarks=remarks||''; sl.updatedAt=sl.cancelledAt;
  if(isAdminRole&&sl.agentId!==u.id) addNotif(d,sl.agentId,`Your Short Leave for ${sl.shortLeaveDate} was cancelled by ${u.name}`, 'sl_cancelled', sl.id);
  save(d); res.json({ok:true});
});

// ─── USERS ────────────────────────────────────────────────────────────
app.get('/api/users', isAdm, (_,r)=>{
  r.json(load().users.map(u=>({id:u.id,username:u.username,role:u.role,name:u.name,permissions:u.permissions||{},createdAt:u.createdAt})));
});
app.post('/api/users', isAdm, (req,res)=>{
  const d=load(); const {username,password,name,role,permissions}=req.body;
  if(!username||!password||!name) return res.status(400).json({error:'All fields required'});
  if(d.users.find(u=>u.username===username)) return res.status(400).json({error:'Username already taken'});
  const creator=req.session.user;
  if(creator.role!=='superadmin'&&(role==='admin'||role==='superadmin')) return res.status(403).json({error:'Only Super Admin can create Admin accounts'});
  const u={id:Date.now(),username,name,password:bcrypt.hashSync(password,10),role:role||'operator',permissions:permissions||{},createdAt:new Date().toISOString()};
  d.users.push(u); save(d); res.json({ok:true,id:u.id});
});
app.put('/api/users/:id', isAdm, (req,res)=>{
  const d=load(); const i=d.users.findIndex(u=>u.id==req.params.id);
  if(i<0) return res.status(404).json({error:'User not found'});
  const creator=req.session.user;
  const {name,password,role,permissions}=req.body;
  if(name) d.users[i].name=name;
  if(role&&creator.role==='superadmin') d.users[i].role=role;
  if(permissions!==undefined) d.users[i].permissions=permissions;
  if(password) d.users[i].password=bcrypt.hashSync(password,10);
  if(req.session.user.id==req.params.id){req.session.user.name=d.users[i].name;req.session.user.permissions=d.users[i].permissions}
  save(d); res.json({ok:true});
});
app.delete('/api/users/:id', isSA, (req,res)=>{
  const d=load(); const u=d.users.find(x=>x.id==req.params.id);
  if(u?.username==='superadmin') return res.status(400).json({error:'Cannot delete superadmin'});
  d.users=d.users.filter(x=>x.id!=req.params.id); save(d); res.json({ok:true});
});
app.put('/api/users/:id/password', auth, (req,res)=>{
  const d=load(); const u=req.session.user; const tid=+req.params.id;
  if(u.role!=='superadmin'&&u.id!==tid) return res.status(403).json({error:'Cannot change others passwords'});
  const {password}=req.body;
  if(!password||password.length<6) return res.status(400).json({error:'Min 6 characters'});
  const i=d.users.findIndex(x=>x.id===tid);
  if(i<0) return res.status(404).json({error:'User not found'});
  d.users[i].password=bcrypt.hashSync(password,10); save(d); res.json({ok:true});
});

// ─── SETTINGS ─────────────────────────────────────────────────────────
app.get('/api/settings', isAdm, (_,r)=>r.json(load().settings));
app.put('/api/settings', isAdm, (req,res)=>{const d=load();d.settings={...d.settings,...req.body};save(d);res.json({ok:true});});
app.get('/api/access', isAdm, (_,r)=>r.json(load().moduleAccess));
app.put('/api/access', isAdm, (req,res)=>{const d=load();d.moduleAccess=req.body;save(d);res.json({ok:true});});

// ─── BACKUP ───────────────────────────────────────────────────────────
app.post('/api/backup', isAdm, (req,res)=>{
  try{
    const d=load(); const ts=new Date().toISOString().replace(/[:.]/g,'-');
    const fn=`backup_${ts}.json`, fp=path.join(BKDIR,fn);
    fs.writeFileSync(fp,JSON.stringify(d,null,2));
    const bks=fs.readdirSync(BKDIR).filter(f=>f.endsWith('.json')).sort();
    if(bks.length>30) bks.slice(0,bks.length-30).forEach(b=>fs.unlinkSync(path.join(BKDIR,b)));
    res.json({ok:true,file:fn,size:fs.statSync(fp).size});
  }catch(e){res.status(500).json({error:e.message});}
});
app.get('/api/backups', isAdm, (_,r)=>{
  try{const bks=fs.readdirSync(BKDIR).filter(f=>f.endsWith('.json')).sort().reverse();r.json(bks.map(f=>({name:f,size:fs.statSync(path.join(BKDIR,f)).size,created:fs.statSync(path.join(BKDIR,f)).mtime})));}catch{r.json([]);}
});
app.get('/api/backups/:name/download', isAdm, (req,res)=>{
  const fp=path.join(BKDIR,req.params.name);
  if(!fs.existsSync(fp)||!req.params.name.endsWith('.json')) return res.status(404).json({error:'Not found'});
  res.download(fp);
});
app.post('/api/backups/:name/restore', isSA, (req,res)=>{
  const fp=path.join(BKDIR,req.params.name);
  if(!fs.existsSync(fp)) return res.status(404).json({error:'Not found'});
  try{
    const bk=JSON.parse(fs.readFileSync(fp,'utf8'));
    const ts=new Date().toISOString().replace(/[:.]/g,'-');
    fs.writeFileSync(path.join(BKDIR,`pre_restore_${ts}.json`),fs.readFileSync(DB,'utf8'));
    save(bk); res.json({ok:true});
  }catch(e){res.status(500).json({error:e.message});}
});
app.delete('/api/backups/:name', isSA, (req,res)=>{
  const fp=path.join(BKDIR,req.params.name);
  if(!fs.existsSync(fp)) return res.status(404).json({error:'Not found'});
  fs.unlinkSync(fp); res.json({ok:true});
});
setInterval(()=>{
  try{
    const d=load(); const ts=new Date().toISOString().replace(/[:.]/g,'-');
    fs.writeFileSync(path.join(BKDIR,`auto_${ts}.json`),JSON.stringify(d,null,2));
    const bks=fs.readdirSync(BKDIR).filter(f=>f.endsWith('.json')).sort();
    if(bks.length>30) bks.slice(0,bks.length-30).forEach(b=>fs.unlinkSync(path.join(BKDIR,b)));
  }catch(e){console.error('Auto backup:',e.message);}
},24*60*60*1000);

// ─── SPA ──────────────────────────────────────────────────────────────
app.get('*', (_,r)=>r.sendFile(path.join(__dirname,'public','index.html')));
app.listen(PORT,()=>{
  console.log(`\n✅  C-Serv.AI v3  →  http://localhost:${PORT}`);
  console.log('    superadmin / Super@123\n    admin / Admin@123\n    operator / Op@123\n');
});
