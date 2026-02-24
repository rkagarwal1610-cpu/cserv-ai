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
        role:'superadmin', name:'Super Administrator', createdAt:new Date().toISOString(), permissions:{} },
      { id:2, username:'admin', password:bcrypt.hashSync('Admin@123',10),
        role:'admin', name:'Administrator', createdAt:new Date().toISOString(),
        permissions:{
          roster:{ view:true,generate:true,save:true,export:true,editAgents:true,editRules:true,approve:true },
          shortleave:{ view:true,approve:true,reject:true,cancel:true,dashboard:true }
        }},
      { id:3, username:'operator', password:bcrypt.hashSync('Op@123',10),
        role:'operator', name:'Team Operator', createdAt:new Date().toISOString(),
        permissions:{ roster:{ view:true }, shortleave:{ view:true,apply:true } }}
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
    // agentLeaves: { "AgentName": { "2026-03": { leaves:[dates], fixedWOs:[dates] } } }
    agentLeaves:{},
    rosters:[],
    shortLeaves:[],
    holidays:[],
    notifications:[],
    rules:{
      targetWOBase:8, extraWOIfFifthSunday:true,
      maxAgentsOnWOPerDay:2, maxConsecutiveWO:3, maxConsecExtraWOPerAgent:1,
      pairingRule:'senior_junior', minL0Working:true,
      holidayMaxWO:1, holidayMaxAgents:2, allowL0OnHolidayWO:true,
      shortLeaveMonthlyLimit:3
    },
    rosterFormat:{
      showEmp:true,showLevel:true,showRole:true,showDept:false,
      cellWO:'WO',cellROI:'ROI',cellHoliday:'H'
    },
    settings:{ appName:'C-Serv.AI', orgName:'Customer Service Team' },
    moduleAccess:{ roster:true, shortleave:true }
  };
  save(d); return d;
}

app.use(bp.json({ limit:'10mb' }));
app.use(bp.urlencoded({ extended:true }));
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({
  secret: process.env.SESSION_SECRET || 'cservai-v4-key-2026',
  resave:false, saveUninitialized:false,
  cookie:{ maxAge:8*60*60*1000, httpOnly:true }
}));

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

// ── NOTIFICATIONS ─────────────────────────────────────────────────────
function addNotif(d, toId, msg, type='info', refId=null) {
  if(!d.notifications) d.notifications=[];
  d.notifications.unshift({ id:Date.now()+Math.random(), toId, msg, type, refId, read:false, createdAt:new Date().toISOString() });
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
  if(n) n.read=true; save(d); res.json({ok:true});
});

// ── AUTH ──────────────────────────────────────────────────────────────
app.post('/api/register', (req,res) => {
  const d=load(); const {username,password,name}=req.body;
  if(!username||!password||!name) return res.status(400).json({error:'All fields required'});
  if(password.length<6) return res.status(400).json({error:'Password must be 6+ characters'});
  if(d.users.find(u=>u.username===username)) return res.status(400).json({error:'Username already taken'});
  const u={id:Date.now(),username,name,password:bcrypt.hashSync(password,10),role:'operator',
    permissions:{ roster:{view:true}, shortleave:{view:true,apply:true} },
    createdAt:new Date().toISOString()};
  d.users.push(u);
  d.users.filter(u2=>['superadmin','admin'].includes(u2.role))
    .forEach(u2=>addNotif(d,u2.id,`New user registered: ${name} (${username})`,'user',u.id));
  save(d); res.json({ok:true,id:u.id});
});
app.post('/api/login', (req,res) => {
  const {username,password}=req.body; const d=load();
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

// ── AGENTS ────────────────────────────────────────────────────────────
app.get('/api/agents', auth, (_,r)=>r.json(load().agents));
app.put('/api/agents', perm('roster','editAgents'), (req,res)=>{
  const d=load(); d.agents=req.body; save(d); res.json({ok:true});
});

// ── AGENT LEAVES (per-agent leave dates & fixed WO dates) ─────────────
// GET all agent leaves for a month
app.get('/api/agentleaves', isAdm, (req,res)=>{
  const d=load(); res.json(d.agentLeaves||{});
});
// PUT leaves for one agent+month: { agentName, monthKey:"2026-03", leaves:["2026-03-05",...], fixedWOs:["2026-03-10",...] }
app.put('/api/agentleaves', isAdm, (req,res)=>{
  const d=load(); const {agentName,monthKey,leaves,fixedWOs}=req.body;
  if(!d.agentLeaves) d.agentLeaves={};
  if(!d.agentLeaves[agentName]) d.agentLeaves[agentName]={};
  d.agentLeaves[agentName][monthKey]={leaves:leaves||[],fixedWOs:fixedWOs||[]};
  save(d); res.json({ok:true});
});

// ── HOLIDAYS ──────────────────────────────────────────────────────────
app.get('/api/holidays', auth, (_,r)=>r.json(load().holidays));
app.put('/api/holidays', isAdm, (req,res)=>{
  const d=load(); d.holidays=req.body; save(d); res.json({ok:true});
});

// ── RULES ─────────────────────────────────────────────────────────────
app.get('/api/rules', auth, (_,r)=>r.json(load().rules));
app.put('/api/rules', isAdm, (req,res)=>{
  const d=load(); d.rules={...d.rules,...req.body}; save(d); res.json({ok:true});
});

// ── ROSTER FORMAT ─────────────────────────────────────────────────────
app.get('/api/rosterformat', auth, (_,r)=>r.json(load().rosterFormat));
app.put('/api/rosterformat', isAdm, (req,res)=>{
  const d=load(); d.rosterFormat={...d.rosterFormat,...req.body}; save(d); res.json({ok:true});
});

// ── ROSTERS ───────────────────────────────────────────────────────────
app.get('/api/rosters', auth, (req,res)=>{
  const d=load(); const u=req.session.user;
  const list=['superadmin','admin'].includes(u.role)?d.rosters:d.rosters.filter(r=>r.approved);
  res.json(list.map(r=>({id:r.id,title:r.title,month:r.month,year:r.year,agentCount:r.agentCount,
    targetWO:r.targetWO,savedAt:r.savedAt,savedBy:r.savedBy,
    approved:r.approved||false,approvedAt:r.approvedAt||null,approvedBy:r.approvedBy||null})));
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
app.put('/api/rosters/:id/approve', perm('roster','approve'), (req,res)=>{
  const d=load(); const u=req.session.user;
  const r=d.rosters.find(x=>x.id==req.params.id);
  if(!r) return res.status(404).json({error:'Not found'});
  r.approved=true; r.approvedAt=new Date().toISOString(); r.approvedBy=u.name;
  d.users.filter(u2=>u2.role==='operator')
    .forEach(u2=>addNotif(d,u2.id,`Roster "${r.title}" has been approved and is now visible.`,'roster',r.id));
  save(d); res.json({ok:true});
});
app.delete('/api/rosters/:id', isAdm, (req,res)=>{
  const d=load(); d.rosters=d.rosters.filter(x=>x.id!=req.params.id); save(d); res.json({ok:true});
});

// ── XLSX EXPORT ───────────────────────────────────────────────────────
app.post('/api/rosters/export/xlsx', perm('roster','export'), async (req,res)=>{
  try{
    const {rosterId}=req.body;
    const d=load();
    const rData=d.rosters.find(x=>x.id==rosterId);
    if(!rData) return res.status(404).json({error:'Not found'});

    const MN=['January','February','March','April','May','June','July','August','September','October','November','December'];
    const DS=['Su','Mo','Tu','We','Th','Fr','Sa'];

    // Rebuild schedule & day-type maps from serialised data
    const dow={};(rData.dowSer||[]).forEach(([k,v])=>dow[+k]=v);
    const dt={};(rData.dtypeSer||[]).forEach(([k,v])=>dt[+k]=v);
    const schedMap={};(rData.schedSer||[]).forEach(([n,v])=>schedMap[n]=v);

    // Role label helper (same logic as frontend)
    const roleLabel=ag=>{
      const l=ag.level;
      if(l===0)return'TL';
      if(l<=3)return`Sr.L${l}`;
      if(l<=6)return`Jr.L${l}`;
      return'Trainee';
    };

    const wb=new ExcelJS.Workbook();
    wb.creator='C-Serv.AI'; wb.created=new Date();
    const ws=wb.addWorksheet(`${MN[rData.month]} ${rData.year}`);

    // ── Column definitions ─────────────────────────────────────────────
    // Fixed: Emp#(A), Name(B), Level(C), Role(D), Total WO(E)
    // Then day columns: day 1-N, then extra "Su" summary col at end
    const days=rData.days||[];
    const totalCols=5+days.length+1; // fixed(5) + days + summary

    // Set column widths
    ws.getColumn(1).width=10;   // Emp#
    ws.getColumn(2).width=22;   // Name
    ws.getColumn(3).width=7;    // Level
    ws.getColumn(4).width=14;   // Role
    ws.getColumn(5).width=9;    // Total WO
    for(let i=6;i<=5+days.length;i++) ws.getColumn(i).width=5;
    ws.getColumn(5+days.length+1).width=6; // Su summary col

    // ── ROW 1: Header row ──────────────────────────────────────────────
    // Emp#, Name, Level, Role, Total WO, 1, 2, 3 ... N, Su
    const hdrValues=['Emp#','Name','Level','Role','Total WO',
      ...days.map(d=>String(d)),
      'Su'
    ];
    const hRow=ws.getRow(1);
    hRow.height=20;
    hdrValues.forEach((v,i)=>{
      const cell=hRow.getCell(i+1);
      cell.value=v;
      cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF0D1526'}};
      cell.font={bold:true,color:{argb:'FF00D4FF'},size:9};
      cell.alignment={horizontal:'center',vertical:'middle',wrapText:false};
    });

    // ── ROW 2: Day-of-week numbers row ────────────────────────────────
    // Cols A-D empty, Col E = 1, Col F = 2 (week day numbers 1=Sun..7=Sat)
    // Matching the sample: row2 has blank A-D, then 1,2,3... for weekday number
    const r2=ws.getRow(2);
    r2.height=14;
    // A-E blank
    [1,2,3,4].forEach(i=>{r2.getCell(i).value='';});
    // Day-of-week numbers: the sample shows 1,2,3... which are the dow numbers
    // (1=Sun, 2=Mon, 3=Tue, 4=Wed, 5=Thu, 6=Fri, 7=Sat in Excel convention)
    // From the sample: day1 is Sun, col E2=1 (Sun=1 in 1-indexed)
    days.forEach((d,i)=>{
      const dowVal=dow[d]; // 0=Sun,1=Mon...6=Sat (JS)
      const xlDow=dowVal===0?1:dowVal+1; // convert to Excel 1-indexed (Sun=1)
      const cell=r2.getCell(5+i+1); // starts at col F (6)
      cell.value=xlDow;
      // Color Saturdays amber, Sundays red, rest default
      if(dowVal===0){cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF2D0A0F'}};cell.font={color:{argb:'FFffb830'},size:8}}
      else if(dowVal===6){cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF2D1F00'}};cell.font={color:{argb:'FFf59e0b'},size:8}}
      else{cell.font={color:{argb:'FF4a6080'},size:8}}
      cell.alignment={horizontal:'center',vertical:'middle'};
    });
    // Last "Su" col row2 blank
    r2.getCell(5+days.length+1).value='';

    // Color Saturday header cols in row 1 too
    days.forEach((d,i)=>{
      const dowVal=dow[d];
      const hCell=hRow.getCell(6+i);
      if(dowVal===0){// Sunday
        hCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF2D0A0F'}};
        hCell.font={bold:true,color:{argb:'FFffb830'},size:9};
      } else if(dowVal===6){// Saturday
        hCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF2D1F00'}};
        hCell.font={bold:true,color:{argb:'FFf59e0b'},size:9};
      } else if(dt[d]==='HOL'){
        hCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF0D2D1E'}};
        hCell.font={bold:true,color:{argb:'FF00e5a0'},size:9};
      }
    });

    // ── Agent rows (Row 3+) ───────────────────────────────────────────
    rData.agents.forEach((ag,ri)=>{
      const sc=schedMap[ag.name]||{};
      const row=ws.addRow([]);
      row.height=18;
      const rowIdx=ri+3; // rows start at 3

      // Agent info cols A-D: dark navy bg
      const agBg='FF0D1526';
      [[1,ag.emp],[2,ag.name],[3,`L${ag.level}`],[4,roleLabel(ag)]].forEach(([col,val])=>{
        const c=row.getCell(col);
        c.value=val;
        c.fill={type:'pattern',pattern:'solid',fgColor:{argb:agBg}};
        c.font={color:{argb:'FFbccde0'},size:9};
        c.alignment={vertical:'middle'};
        if(col===1||col===3) c.alignment={horizontal:'center',vertical:'middle'};
      });

      // Total WO col (E): count only WO days (not HOL)
      let woCount=0;
      days.forEach(d=>{
        const s=sc[d]||'ROI';
        if(s==='WO') woCount++;
      });
      const twoCell=row.getCell(5);
      twoCell.value=woCount;
      twoCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF2D1F00'}};
      twoCell.font={bold:true,color:{argb:'FFffb830'},size:9};
      twoCell.alignment={horizontal:'center',vertical:'middle'};

      // Day cells
      days.forEach((d,ci)=>{
        const s=sc[d]||'ROI';
        const t=dt[d]||'WORK';
        const dowVal=dow[d]; // 0=Sun,6=Sat
        const cell=row.getCell(6+ci);
        cell.alignment={horizontal:'center',vertical:'middle'};
        cell.font={bold:true,size:8};

        // Determine cell value and color:
        // Sunday     → WO, red tint  FF2D0A0F / FFffb830
        // Saturday WO → WO, amber tint FF2D1F00 / FFf59e0b
        // Saturday ROI (agent works) → ROI, amber tint FF2D1F00 / FF888
        // Holiday    → H, green tint FF0D2D1E / FF00e5a0
        // Extra WO   → WO, purple tint FF1A1540 / FFa78bfa
        // Leave      → LV, rose FF2D0A0F / FFff3d5a
        // ROI        → ROI, dark bg / dim text

        if(dowVal===0){// Sunday
          cell.value='WO';
          cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF2D0A0F'}};
          cell.font={bold:true,color:{argb:'FFffb830'},size:8};
        } else if(t==='HOL'){// Holiday
          cell.value='H';
          cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF0D2D1E'}};
          cell.font={bold:true,color:{argb:'FF00e5a0'},size:8};
        } else if(dowVal===6){// Saturday
          if(s==='WO'){
            cell.value='WO';
            cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF2D1F00'}};
            cell.font={bold:true,color:{argb:'FFf59e0b'},size:8};
          } else {
            cell.value='ROI';
            cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF2D1F00'}};
            cell.font={bold:false,color:{argb:'FF8a7040'},size:8};
          }
        } else if(s==='LV'){// Leave
          cell.value='LV';
          cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF2D0A0F'}};
          cell.font={bold:true,color:{argb:'FFff3d5a'},size:8};
        } else if(s==='WO'){// Extra weekday WO
          cell.value='WO';
          cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1A1540'}};
          cell.font={bold:true,color:{argb:'FFa78bfa'},size:8};
        } else {// ROI
          cell.value='ROI';
          cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:ri%2===0?'FF060E1A':'FF040A14'}};
          cell.font={bold:false,color:{argb:'FF4a6080'},size:8};
        }
      });

      // Last "Su" summary col - blank with agent bg
      const lastCell=row.getCell(6+days.length);
      lastCell.value='';
      lastCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:agBg}};
    });

    // Freeze panes: freeze first 4 cols (A-D) and header row
    ws.views=[{state:'frozen',xSplit:5,ySplit:1}];

    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename="Roster_${MN[rData.month]}_${rData.year}.xlsx"`);
    await wb.xlsx.write(res); res.end();
  }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

// ── XLSX INLINE EXPORT (Shifts & Weekly Offs Import format) ───────────
app.post('/api/rosters/export/xlsx/inline', perm('roster','export'), async (req,res)=>{
  try{
    const r=req.body;
    const MN2=['January','February','March','April','May','June','July','August','September','October','November','December'];
    const yr=r.yr, mo=r.mo, agents=r.agents||[], days=r.days||[], sc=r.sc||{};
    const dow=r.dow||{};
    const agSatWO=r.agSatWO||{};
    const holDays=new Set(r.holDays||[]);
    const vLabel=r.versionLabel||'';

    const wb=new ExcelJS.Workbook();
    wb.creator='C-Serv.AI';wb.created=new Date();
    const ws=wb.addWorksheet('Shifts & Weekly Offs Import');

    // Fills & fonts
    const fHdr ={type:'pattern',pattern:'solid',fgColor:{argb:'FFD3D3D3'}};
    const fTop ={type:'pattern',pattern:'solid',fgColor:{argb:'FFADD8E6'}};
    const fWO  ={type:'pattern',pattern:'solid',fgColor:{argb:'FFCCC0D9'}};
    const fHol ={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFC000'}};
    const fLv  ={type:'pattern',pattern:'solid',fgColor:{argb:'FF92D050'}};
    const fNone={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFFFF'}};
    const fSunH={type:'pattern',pattern:'solid',fgColor:{argb:'FFFF9999'}};
    const fSatH={type:'pattern',pattern:'solid',fgColor:{argb:'FFDDD0EA'}};
    const fCnt ={type:'pattern',pattern:'solid',fgColor:{argb:'FFE2EFDA'}};
    const thinBdr={style:'thin',color:{argb:'FF999999'}};
    const bdr={top:thinBdr,bottom:thinBdr,left:thinBdr,right:thinBdr};
    const center={horizontal:'center',vertical:'middle'};
    const left={horizontal:'left',vertical:'middle'};

    // Column widths
    ws.getColumn(1).width=19; ws.getColumn(2).width=29;
    ws.getColumn(3).width=18; ws.getColumn(4).width=12;
    for(let i=5;i<=4+days.length;i++) ws.getColumn(i).width=12.9;
    ws.getColumn(5+days.length).width=10;

    const nd=days.length;

    // Row 1: light blue bar
    const r1=ws.getRow(1); r1.height=14;
    for(let c=1;c<=5+nd;c++){
      const cell=r1.getCell(c);
      cell.fill=fTop; cell.border=bdr;
    }
    r1.getCell(1).value='  ';

    // Row 2: grey headers + date columns
    const r2=ws.getRow(2); r2.height=32;
    [['Employee Number',1],['Employee Name',2],['Department',3],['Location',4]].forEach(([h,c])=>{
      const cell=r2.getCell(c);
      cell.value=h; cell.fill=fHdr; cell.border=bdr;
      cell.font={bold:true,size:9};
      cell.alignment=left;
    });
    days.forEach((d,i)=>{
      const c=5+i; const dw=+dow[d];
      const cell=r2.getCell(c);
      cell.value=new Date(yr,mo,d);
      cell.numFmt='dd/mmm/yyyy';
      cell.fill=dw===0?fSunH:(dw===6?fSatH:fHdr);
      cell.font={bold:true,size:9};
      cell.alignment=left;
      cell.border=bdr;
    });
    const twHdr=r2.getCell(5+nd);
    twHdr.value='Total WO'; twHdr.fill=fHdr; twHdr.border=bdr;
    twHdr.font={bold:true,size:9}; twHdr.alignment=center;

    // Agent rows
    agents.forEach((ag,ri)=>{
      const rowIdx=3+ri; const agSc=sc[ag.name]||{};
      const row=ws.getRow(rowIdx); row.height=16;
      [[1,ag.emp],[2,ag.name],[3,ag.dept||''],[4,ag.loc||'']].forEach(([c,v])=>{
        const cell=row.getCell(c);
        cell.value=v; cell.fill=fNone; cell.border=bdr;
        cell.font={size:9}; cell.alignment=c<=2?left:center;
      });
      days.forEach((d,i)=>{
        const c=5+i; const v=agSc[d]||'ROI'; const dw=+dow[d];
        const cell=row.getCell(c);
        cell.border=bdr; cell.font={size:9,bold:v==='WO'};
        cell.alignment=center;
        if(v==='WO'){cell.value='WO';cell.fill=fWO;}
        else if(v==='HOL'||v==='Holiday'){cell.value='Holiday';cell.fill=fHol;}
        else if(v==='LV'||v==='Leave'){cell.value='Leave';cell.fill=fLv;}
        else{cell.value='ROI';cell.fill=fNone;}
      });
      // COUNTIF Total WO formula
      const cs=ws.getColumn(5).letter; const ce=ws.getColumn(4+nd).letter;
      const twCell=row.getCell(5+nd);
      twCell.value={formula:`COUNTIF(${cs}${rowIdx}:${ce}${rowIdx},"WO")`};
      twCell.fill=fNone; twCell.border=bdr;
      twCell.font={bold:true,size:9}; twCell.alignment=center;
    });

    // Count row
    const crIdx=3+agents.length;
    const cr=ws.getRow(crIdx); cr.height=14;
    days.forEach((d,i)=>{
      const c=5+i; const col=ws.getColumn(c).letter;
      const cell=cr.getCell(c);
      cell.value={formula:`COUNTIF(${col}3:${col}${2+agents.length},"ROI")`};
      cell.fill=fCnt; cell.border=bdr;
      cell.font={size:8,italic:true}; cell.alignment=center;
    });

    // Legend
    const legends=[[crIdx+3,'WO',fWO],[crIdx+4,'Leave',fLv],[crIdx+5,'Sunday and Holiday Working',fTop],[crIdx+6,'Holiday',fHol]];
    legends.forEach(([lr,txt,fill])=>{
      const cell=ws.getRow(lr).getCell(2);
      cell.value=txt; cell.fill=fill;
      cell.font={size:9}; cell.alignment=left;
    });

    ws.views=[{state:'frozen',xSplit:4,ySplit:2}];

    const fn=`Roster_${MN2[mo]}_${yr}${vLabel?'_'+vLabel:''}.xlsx`;
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename="${fn}"`);
    await wb.xlsx.write(res); res.end();
  }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

// ── SHORT LEAVE ───────────────────────────────────────────────────────
app.get('/api/shortleaves', auth, (req,res)=>{
  const d=load(); const u=req.session.user;
  const list=['superadmin','admin'].includes(u.role)?d.shortLeaves:d.shortLeaves.filter(sl=>sl.agentId===u.id);
  res.json(list);
});
app.post('/api/shortleaves', perm('shortleave','apply'), (req,res)=>{
  const d=load(); const u=req.session.user;
  const {shortLeaveDate,halfDay,reason,requestDate}=req.body;
  if(!shortLeaveDate) return res.status(400).json({error:'Short leave date required'});
  if(!halfDay||!['1st Half','2nd Half'].includes(halfDay)) return res.status(400).json({error:'Select 1st Half or 2nd Half'});
  const rd=requestDate||new Date().toISOString().slice(0,10);
  const slMonth=shortLeaveDate.slice(0,7);
  const monthCount=d.shortLeaves.filter(sl=>sl.agentId===u.id&&sl.shortLeaveDate.slice(0,7)===slMonth&&sl.status!=='Cancelled').length;
  if(monthCount>=(d.rules.shortLeaveMonthlyLimit||3)) return res.status(400).json({error:`Monthly limit reached for ${slMonth}`});
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
  d.users.filter(u2=>['superadmin','admin'].includes(u2.role))
    .forEach(u2=>addNotif(d,u2.id,`${u.name} applied Short Leave for ${shortLeaveDate} (${halfDay})`,'sl_new',sl.id));
  save(d); res.json({ok:true,id:sl.id});
});
app.put('/api/shortleaves/:id/approve', perm('shortleave','approve'), (req,res)=>{
  const d=load(); const u=req.session.user; const {remarks}=req.body;
  const sl=d.shortLeaves.find(x=>x.id==req.params.id);
  if(!sl) return res.status(404).json({error:'Not found'});
  if(sl.status!=='Pending') return res.status(400).json({error:'Only pending requests can be approved'});
  sl.status='Approved'; sl.approvedAt=new Date().toISOString(); sl.approvedBy=u.name;
  sl.remarks=remarks||''; sl.updatedAt=sl.approvedAt;
  addNotif(d,sl.agentId,`Your Short Leave for ${sl.shortLeaveDate} (${sl.halfDay}) has been Approved by ${u.name}`,'sl_approved',sl.id);
  save(d); res.json({ok:true});
});
app.put('/api/shortleaves/:id/reject', perm('shortleave','reject'), (req,res)=>{
  const d=load(); const u=req.session.user; const {remarks}=req.body;
  const sl=d.shortLeaves.find(x=>x.id==req.params.id);
  if(!sl) return res.status(404).json({error:'Not found'});
  if(sl.status!=='Pending') return res.status(400).json({error:'Only pending requests can be rejected'});
  sl.status='Rejected'; sl.rejectedAt=new Date().toISOString(); sl.rejectedBy=u.name;
  sl.remarks=remarks||''; sl.updatedAt=sl.rejectedAt;
  addNotif(d,sl.agentId,`Your Short Leave for ${sl.shortLeaveDate} (${sl.halfDay}) was Rejected. Reason: ${remarks||'—'}`,'sl_rejected',sl.id);
  save(d); res.json({ok:true});
});
app.put('/api/shortleaves/:id/cancel', auth, (req,res)=>{
  const d=load(); const u=req.session.user; const {remarks}=req.body;
  const sl=d.shortLeaves.find(x=>x.id==req.params.id);
  if(!sl) return res.status(404).json({error:'Not found'});
  const isA=['superadmin','admin'].includes(u.role);
  if(!isA&&sl.agentId!==u.id) return res.status(403).json({error:'Cannot cancel others requests'});
  if(!isA&&sl.status!=='Pending') return res.status(400).json({error:'You can only cancel pending requests'});
  if(isA&&!['Pending','Approved'].includes(sl.status)) return res.status(400).json({error:'Can only cancel Pending or Approved'});
  sl.status='Cancelled'; sl.cancelledAt=new Date().toISOString(); sl.cancelledBy=u.name;
  sl.remarks=remarks||''; sl.updatedAt=sl.cancelledAt;
  if(isA&&sl.agentId!==u.id) addNotif(d,sl.agentId,`Your Short Leave for ${sl.shortLeaveDate} was cancelled by ${u.name}`,'sl_cancelled',sl.id);
  save(d); res.json({ok:true});
});

// ── USERS ─────────────────────────────────────────────────────────────
app.get('/api/users', isAdm, (_,r)=>{
  r.json(load().users.map(u=>({id:u.id,username:u.username,role:u.role,name:u.name,permissions:u.permissions||{},createdAt:u.createdAt})));
});
app.post('/api/users', isAdm, (req,res)=>{
  const d=load(); const {username,password,name,role,permissions}=req.body;
  if(!username||!password||!name) return res.status(400).json({error:'All fields required'});
  if(d.users.find(u=>u.username===username)) return res.status(400).json({error:'Username taken'});
  const creator=req.session.user;
  if(creator.role!=='superadmin'&&(role==='admin'||role==='superadmin')) return res.status(403).json({error:'Only SA can create admins'});
  const u={id:Date.now(),username,name,password:bcrypt.hashSync(password,10),role:role||'operator',permissions:permissions||{},createdAt:new Date().toISOString()};
  d.users.push(u); save(d); res.json({ok:true,id:u.id});
});
app.put('/api/users/:id', isAdm, (req,res)=>{
  const d=load(); const i=d.users.findIndex(u=>u.id==req.params.id);
  if(i<0) return res.status(404).json({error:'Not found'});
  const creator=req.session.user; const {name,password,role,permissions}=req.body;
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
  if(u.role!=='superadmin'&&u.id!==tid) return res.status(403).json({error:'Cannot change others password'});
  const {password}=req.body;
  if(!password||password.length<6) return res.status(400).json({error:'Min 6 chars'});
  const i=d.users.findIndex(x=>x.id===tid);
  if(i<0) return res.status(404).json({error:'Not found'});
  d.users[i].password=bcrypt.hashSync(password,10); save(d); res.json({ok:true});
});

// ── SETTINGS ─────────────────────────────────────────────────────────
app.get('/api/settings', isAdm, (_,r)=>r.json(load().settings));
app.put('/api/settings', isAdm, (req,res)=>{const d=load();d.settings={...d.settings,...req.body};save(d);res.json({ok:true});});
app.get('/api/access', isAdm, (_,r)=>r.json(load().moduleAccess));
app.put('/api/access', isAdm, (req,res)=>{const d=load();d.moduleAccess=req.body;save(d);res.json({ok:true});});

// ── BACKUP ────────────────────────────────────────────────────────────
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

app.get('*', (_,r)=>r.sendFile(path.join(__dirname,'public','index.html')));
app.listen(PORT,()=>{
  console.log(`\n✅  C-Serv.AI v4  →  http://localhost:${PORT}`);
  console.log('    superadmin / Super@123 | admin / Admin@123 | operator / Op@123\n');
});
