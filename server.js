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
      {emp:'EP0200',satConfig:'1st3rd',name:'Shrikant Nayak',level:0,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0269',satConfig:'2nd4th',name:'Ritu Singh',level:0,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0505',satConfig:'2nd4th',name:'Rohit Kumar Agarwal',level:0,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0563',satConfig:'1st3rd',name:'Himanshi Khowal',level:2,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0564',satConfig:'1st3rd',name:'Chetan Goel',level:1,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0523',satConfig:'1st3rd',name:'Sushant Kumar Suman',level:2,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0560',satConfig:'1st3rd',name:'Mohit Singh',level:3,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0678',satConfig:'1st3rd',name:'Abhay Pratap',level:4,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0726',satConfig:'1st3rd',name:'Swagata Bhoumik',level:6,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0848',satConfig:'1st3rd',name:'Deepak Gupta',level:7,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0442',satConfig:'2nd4th',name:'Shivam Garg',level:1,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0524',satConfig:'2nd4th',name:'Anurag Tiwari',level:4,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0567',satConfig:'2nd4th',name:'Triloki Varshney',level:1,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0557',satConfig:'2nd4th',name:'Sujit Kumar',level:3,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0741',satConfig:'2nd4th',name:'Amarnath Vishwakarma',level:5,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0673',satConfig:'2nd4th',name:'Dhruv Mishra',level:5,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0798',satConfig:'2nd4th',name:'Naveen Kumar S',level:6,dept:'Customer Service',loc:'Gurgaon-US'}
    ],
    // agentLeaves: { "AgentName": { "2026-03": { leaves:[dates], fixedWOs:[dates] } } }
    agentLeaves:{},
    rosters:[],
    shortLeaves:[],
    holidays:[
      {name:'Republic Day',        date:'2026-01-26'},
      {name:'Holi',                date:'2026-03-04'},
      {name:'Ram Navami',          date:'2026-03-27'},
      {name:'Good Friday',         date:'2026-04-03'},
      {name:'Ambedkar Jayanti',    date:'2026-04-14'},
      {name:'Maharashtra Day',     date:'2026-05-01'},
      {name:'Buddha Purnima',      date:'2026-05-12'},
      {name:'Eid ul-Adha',         date:'2026-06-17'},
      {name:'Muharram',            date:'2026-07-06'},
      {name:'Independence Day',    date:'2026-08-15'},
      {name:'Janmashtami',         date:'2026-08-23'},
      {name:'Gandhi Jayanti',      date:'2026-10-02'},
      {name:'Dussehra',            date:'2026-10-09'},
      {name:'Diwali',              date:'2026-10-28'},
      {name:'Guru Nanak Jayanti',  date:'2026-11-14'},
      {name:'Christmas',           date:'2026-12-25'},
    ],
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

// ── SECURITY HEADERS ──────────────────────────────────────────────────
app.use((req,res,next)=>{
  res.setHeader('X-Content-Type-Options','nosniff');
  res.setHeader('X-Frame-Options','DENY');
  res.setHeader('X-XSS-Protection','1; mode=block');
  res.setHeader('Referrer-Policy','strict-origin-when-cross-origin');
  res.setHeader('Permissions-Policy','camera=(),microphone=(),geolocation=()');
  res.setHeader('Content-Security-Policy',
    "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src 'self' https://fonts.gstatic.com; img-src 'self' data:; connect-src 'self'"
  );
  if(process.env.NODE_ENV==='production'){
    res.setHeader('Strict-Transport-Security','max-age=31536000; includeSubDomains');
  }
  next();
});

// ── COMPRESSION (gzip for text responses) ─────────────────────────────
app.use((req,res,next)=>{
  const ae=req.headers['accept-encoding']||'';
  if(!ae.includes('gzip')){return next();}
  const {Transform}=require('stream');
  const zlib=require('zlib');
  // Only compress API JSON and HTML — skip binary (xlsx)
  const origSend=res.send.bind(res);
  const origJson=res.json.bind(res);
  res.json=function(body){
    const str=JSON.stringify(body);
    if(str.length<1024){res.setHeader('Content-Type','application/json');return origSend(str);}
    const buf=Buffer.from(str);
    zlib.gzip(buf,(err,gz)=>{
      if(err){res.setHeader('Content-Type','application/json');return origSend(str);}
      res.setHeader('Content-Type','application/json');
      res.setHeader('Content-Encoding','gzip');
      res.setHeader('Vary','Accept-Encoding');
      origSend(gz);
    });
  };
  next();
});

// ── RATE LIMITING ──────────────────────────────────────────────────────
const _rl={};
function rateLimit(windowMs,max,keyFn){
  return (req,res,next)=>{
    const key=keyFn(req);const now=Date.now();
    if(!_rl[key])_rl[key]={count:0,reset:now+windowMs};
    if(now>_rl[key].reset){_rl[key]={count:0,reset:now+windowMs};}
    _rl[key].count++;
    if(_rl[key].count>max){
      const wait=Math.ceil((_rl[key].reset-now)/1000);
      return res.status(429).json({error:`Too many requests. Try again in ${wait}s.`});
    }
    next();
  };
}
// Login: max 10 attempts per IP per 15 minutes
const loginRL=rateLimit(15*60*1000,10,req=>req.ip||req.connection.remoteAddress);
// General API: max 300 per IP per minute
const apiRL=rateLimit(60*1000,300,req=>req.ip||req.connection.remoteAddress);
app.use('/api',apiRL);
// Clean up old RL entries every 30 min
setInterval(()=>{const now=Date.now();Object.keys(_rl).forEach(k=>{if(_rl[k].reset<now)delete _rl[k];});},30*60*1000);

// ── BRUTE-FORCE LOGIN LOCKOUT ──────────────────────────────────────────
const _lk={};// { ip: { fails:N, lockedUntil:ts } }
function loginGuard(req,res,next){
  const ip=req.ip||req.connection.remoteAddress;
  const now=Date.now();
  if(!_lk[ip])_lk[ip]={fails:0,lockedUntil:0};
  if(now<_lk[ip].lockedUntil){
    const wait=Math.ceil((_lk[ip].lockedUntil-now)/1000);
    return res.status(429).json({error:`Account locked. Try again in ${wait}s.`});
  }
  next();
}
function loginFail(req){
  const ip=req.ip||req.connection.remoteAddress;const now=Date.now();
  if(!_lk[ip])_lk[ip]={fails:0,lockedUntil:0};
  _lk[ip].fails++;
  if(_lk[ip].fails>=5)_lk[ip].lockedUntil=now+15*60*1000; // lock 15 min after 5 fails
}
function loginOk(req){
  const ip=req.ip||req.connection.remoteAddress;
  _lk[ip]={fails:0,lockedUntil:0};
}

// ── BODY PARSER + STATIC ───────────────────────────────────────────────
app.use(bp.json({ limit:'2mb' }));  // reduced from 10mb — sufficient for roster data
app.use(bp.urlencoded({ extended:true, limit:'1mb' }));

// Static files with cache headers
app.use(express.static(path.join(__dirname, 'public'),{
  maxAge:'1d',          // cache static assets 1 day
  etag:true,
  lastModified:true,
  setHeaders:(res,filePath)=>{
    // No-cache for index.html so app updates are picked up immediately
    if(filePath.endsWith('index.html')){
      res.setHeader('Cache-Control','no-cache, no-store, must-revalidate');
    }
  }
}));

app.use(session({
  secret: process.env.SESSION_SECRET || (process.env.NODE_ENV==='production' ? (()=>{throw new Error('SESSION_SECRET env var is required in production');})() : 'cservai-dev-secret-change-in-prod-2026-xK9mP'),
  resave:false,
  saveUninitialized:false,
  name:'cservai.sid',   // don't reveal stack via default 'connect.sid'
  cookie:{
    maxAge:8*60*60*1000,
    httpOnly:true,
    secure:process.env.NODE_ENV==='production',   // HTTPS-only in prod
    sameSite:'strict'   // CSRF protection
  }
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
// Register endpoint removed — users are created by superadmin/admin in the UI
app.post('/api/login', loginRL, loginGuard, (req,res) => {
  const {username,password}=req.body;
  // Input validation
  if(!username||!password||typeof username!=='string'||typeof password!=='string'){
    return res.status(400).json({error:'Username and password required'});
  }
  if(username.length>64||password.length>128){
    return res.status(400).json({error:'Invalid credentials'});
  }
  const d=load();
  const u=d.users.find(x=>x.username===username.trim().toLowerCase());
  if(!u||!bcrypt.compareSync(password,u.password)){
    loginFail(req);
    // Uniform delay to prevent timing attacks
    return setTimeout(()=>res.status(401).json({error:'Invalid username or password'}),400);
  }
  loginOk(req);
  // Regenerate session ID on login to prevent session fixation
  req.session.regenerate((err)=>{
    if(err) return res.status(500).json({error:'Session error'});
    req.session.user={id:u.id,username:u.username,role:u.role,name:u.name,permissions:u.permissions||{}};
    req.session.loginAt=Date.now();
    res.json({ok:true,user:req.session.user,moduleAccess:d.moduleAccess,settings:d.settings});
  });
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
// PUT leaves for one agent+month: { agentName, monthKey:"2026-03", leaves:[...], fixedWOs:[...], sunHolWork:[...] }
app.put('/api/agentleaves', isAdm, (req,res)=>{
  const d=load(); const {agentName,monthKey,leaves,fixedWOs,sunHolWork}=req.body;
  if(!d.agentLeaves) d.agentLeaves={};
  if(!d.agentLeaves[agentName]) d.agentLeaves[agentName]={};
  d.agentLeaves[agentName][monthKey]={leaves:leaves||[],fixedWOs:fixedWOs||[],sunHolWork:sunHolWork||[]};
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
        } else if(t==='HOL'){// Holiday — ROI text, Orange Accent 6 bg
          cell.value='ROI';
          cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFED7D31'}};
          cell.font={bold:false,color:{argb:'FF000000'},size:8};
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
        } else if(s==='LV'){// Leave — ROI text, Red bg
          cell.value='ROI';
          cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFF0000'}};
          cell.font={bold:false,color:{argb:'FF000000'},size:8};
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
    const yr=r.yr, mo=r.mo, agents=r.agents||[], days=r.days||[], holDays=new Set((r.holDays||[]).map(Number));
    const vLabel=r.versionLabel||'';
    // Normalize dow: keys may be strings after JSON parse, normalize to number keys
    const dowRaw=r.dow||{};
    const dow={};Object.entries(dowRaw).forEach(([k,v])=>dow[+k]=+v);
    // Normalize sc: day keys may be strings
    const scRaw=r.sc||{};
    const sc={};
    Object.entries(scRaw).forEach(([name,dayMap])=>{
      sc[name]={};
      Object.entries(dayMap||{}).forEach(([d,v])=>sc[name][+d]=v);
    });
    const agSatWO=r.agSatWO||{};

    const wb=new ExcelJS.Workbook();
    wb.creator='C-Serv.AI';wb.created=new Date();
    const ws=wb.addWorksheet('Shifts & Weekly Offs Import');

    // Fills & fonts
    const fHdr ={type:'pattern',pattern:'solid',fgColor:{argb:'FFD3D3D3'}};
    const fTop ={type:'pattern',pattern:'solid',fgColor:{argb:'FFADD8E6'}};
    const fWO  ={type:'pattern',pattern:'solid',fgColor:{argb:'FFCCC0D9'}};
    const fHol ={type:'pattern',pattern:'solid',fgColor:{argb:'FFED7D31'}}; // Orange Accent 6 (Excel standard)
    const fLv  ={type:'pattern',pattern:'solid',fgColor:{argb:'FFFF0000'}}; // Red for Leave
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
      // ALL date header cells: same uniform grey bg
      cell.fill=fHdr;
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
        // Normal black font for all info columns
        cell.font={size:9,color:{argb:'FF000000'}};
        cell.alignment=c<=2?left:center;
      });
      days.forEach((d,i)=>{
        const c=5+i; const v=agSc[d]||'ROI'; const dw=+dow[d];
        const cell=row.getCell(c);
        cell.border=bdr; cell.alignment=center;
        // WO=purple bold | Holiday=ROI text + Orange Accent 6 | Leave=ROI text + Red | ROI=white
        const fYel={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFF00'}}; // Yellow for Sun/Hol working
        if(v==='SW'||v==='HW'){
          // Agent working on Sunday/Holiday — yellow background
          cell.value='ROI'; cell.fill=fYel; cell.font={size:9,bold:false,color:{argb:'FF000000'}};
        } else if(v==='WO'){
          cell.value='WO';  cell.fill=fWO;  cell.font={size:9,bold:true, color:{argb:'FF000000'}};
        } else if(v==='HOL'||v==='H'){
          cell.value='ROI'; cell.fill=fHol; cell.font={size:9,bold:false,color:{argb:'FF000000'}};
        } else if(v==='LV'){
          cell.value='ROI'; cell.fill=fLv;  cell.font={size:9,bold:false,color:{argb:'FF000000'}};
        } else {
          cell.value='ROI'; cell.fill=fNone;cell.font={size:9,bold:false,color:{argb:'FF000000'}};
        }
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
    const legends=[[crIdx+3,'WO',fWO],[crIdx+4,'Leave (shown as ROI)',fLv],[crIdx+5,'Sunday/Holiday (shown as ROI)',fHol],[crIdx+6,'Holiday',fHol]];
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
  if(typeof username!=='string'||!/^[a-z0-9._-]{3,32}$/.test(username.trim().toLowerCase())) return res.status(400).json({error:'Username: 3-32 chars, letters/numbers/._- only'});
  if(password.length<8) return res.status(400).json({error:'Password must be 8+ characters'});
  if(role&&!['admin','operator'].includes(role)) return res.status(400).json({error:'Role must be admin or operator'});
  if(d.users.find(u=>u.username===username.trim().toLowerCase())) return res.status(400).json({error:'Username taken'});
  const creator=req.session.user;
  if(creator.role!=='superadmin'&&(role==='admin'||role==='superadmin')) return res.status(403).json({error:'Only SA can create admins'});
  const u={id:Date.now(),username:username.trim().toLowerCase(),name,password:bcrypt.hashSync(password,10),role:role||'operator',permissions:permissions||{},createdAt:new Date().toISOString()};
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


// ══════════════════════════════════════════════════════════════════════════════
// MODULE 3 — PERFORMANCE DASHBOARD
// Replicates finalapp_v23.py logic entirely server-side
// ══════════════════════════════════════════════════════════════════════════════
const multer = require('multer');
const os = require('os');
const perfUpload = multer({
  dest: os.tmpdir(),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20 MB max
  fileFilter: (_req, file, cb) => {
    const ok = /\.(xlsx|xls|csv)$/i.test(file.originalname);
    cb(ok ? null : new Error('Only .xlsx, .xls or .csv files allowed'), ok);
  }
});

// ── Perf config: GET agents (targets, roles, sort) ───────────────────────────
app.get('/api/perf/config', auth, (_req, res) => {
  const d = load();
  res.json(d.perfConfig || getDefaultPerfConfig());
});

// ── Perf config: PUT (save agents) ──────────────────────────────────────────
app.put('/api/perf/config', isAdm, (req, res) => {
  const d = load();
  const { agents } = req.body;
  if (!Array.isArray(agents)) return res.status(400).json({ error: 'agents array required' });
  d.perfConfig = { agents };
  save(d); res.json({ ok: true });
});

function getDefaultPerfConfig() {
  return {
    agents: [
      { name: 'Shrikant Nayak',       role: 'L1', target: 0, sort: 1  },
      { name: 'Ritu Singh',            role: 'L1', target: 0, sort: 2  },
      { name: 'Rohit Kumar Agarwal',   role: 'L1', target: 0, sort: 3  },
      { name: 'Himanshi Khowal',       role: 'L1', target: 0, sort: 4  },
      { name: 'Chetan Goel',           role: 'L1', target: 0, sort: 5  },
      { name: 'Sushant Kumar Suman',   role: 'L1', target: 0, sort: 6  },
      { name: 'Mohit Singh',           role: 'L1', target: 0, sort: 7  },
      { name: 'Abhay Pratap',          role: 'L1', target: 0, sort: 8  },
      { name: 'Swagata Bhoumik',       role: 'L1', target: 0, sort: 9  },
      { name: 'Deepak Gupta',          role: 'L1', target: 0, sort: 10 },
      { name: 'Shivam Garg',           role: 'L1', target: 0, sort: 11 },
      { name: 'Anurag Tiwari',         role: 'L1', target: 0, sort: 12 },
      { name: 'Triloki Nath',          role: 'L1', target: 0, sort: 13 },
      { name: 'Sujit Kumar',           role: 'L1', target: 0, sort: 14 },
      { name: 'Amarnath Yadav',        role: 'L2', target: 0, sort: 15 },
      { name: 'Dhruv Sharma',          role: 'L2', target: 0, sort: 16 },
      { name: 'Naveen Kumar',          role: 'L2', target: 0, sort: 17 },
    ]
  };
}

// ── Perf process: POST multipart Excel/CSV → returns dashboard JSON ──────────
app.post('/api/perf/process', auth, perfUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const { totalInward = 0, closingUnresolved = 0 } = req.body;
  const tmpPath = req.file.path;

  try {
    const result = await processPerfData(
      tmpPath,
      req.file.originalname,
      parseInt(totalInward) || 0,
      parseInt(closingUnresolved) || 0,
      load().perfConfig || getDefaultPerfConfig()
    );
    res.json(result);
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    try { require('fs').unlinkSync(tmpPath); } catch {}
  }
});

async function processPerfData(filePath, fileName, totalInward, closingUnresolved, perfConfig) {
  const ExcelJS = require('exceljs');
  const wb = new ExcelJS.Workbook();

  const isCSV = /\.csv$/i.test(fileName);
  const rows = [];

  if (isCSV) {
    // Parse CSV manually
    const raw = require('fs').readFileSync(filePath, 'utf8');
    const lines = raw.split(/\r?\n/).filter(l => l.trim());
    if (lines.length < 2) throw new Error('CSV file is empty or has no data rows');
    const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
    for (let i = 1; i < lines.length; i++) {
      const vals = splitCSVLine(lines[i]);
      const row = {};
      headers.forEach((h, idx) => { row[h] = (vals[idx] || '').trim().replace(/^"|"$/g, ''); });
      rows.push(row);
    }
  } else {
    await wb.xlsx.readFile(filePath);
    const ws = wb.worksheets[0];
    if (!ws) throw new Error('Excel file has no worksheets');
    const headers = [];
    ws.getRow(1).eachCell({ includeEmpty: true }, (cell, colNum) => {
      headers[colNum - 1] = String(cell.value || '').trim();
    });
    ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
      if (rowNum === 1) return;
      const obj = {};
      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        let v = cell.value;
        if (v && typeof v === 'object' && v.text) v = v.text;
        if (v instanceof Date) v = v.toISOString();
        obj[headers[colNum - 1]] = v != null ? String(v).trim() : '';
      });
      if (Object.values(obj).some(v => v !== '')) rows.push(obj);
    });
  }

  if (!rows.length) throw new Error('No data rows found in file');

  // Column aliasing (Salesforce export format → expected)
  const colMap = {
    'Case: Case Number': 'Ticket Ticket ID',
    'Case: Priority': 'Ticket Priority',
    'Case: Status': 'Ticket Status',
    'Case: Group *': 'Ticket Group name',
    'Time Log: Created Date': 'Clocked date',
  };
  const firstRow = rows[0];
  const renames = {};
  Object.entries(colMap).forEach(([sf, ex]) => { if (sf in firstRow) renames[sf] = ex; });
  if ('Ticket ID' in firstRow && !('Ticket Ticket ID' in firstRow)) renames['Ticket ID'] = 'Ticket Ticket ID';
  if (Object.keys(renames).length) {
    rows.forEach(row => {
      Object.entries(renames).forEach(([from, to]) => {
        if (from in row) { row[to] = row[from]; delete row[from]; }
      });
    });
  }

  const required = ['Agent', 'Ticket Ticket ID', 'Ticket Group name', 'Ticket Status', 'Ticket Priority', 'Clocked date'];
  for (const col of required) {
    if (!(col in (rows[0] || {}))) throw new Error('Missing required column: "' + col + '"');
  }

  // Normalise & sort rows by TicketID + Agent (dedup logic from Python)
  rows.forEach(r => { r.Agent = (r.Agent || '').trim(); });
  rows.sort((a, b) => {
    const t = String(a['Ticket Ticket ID']).localeCompare(String(b['Ticket Ticket ID']));
    return t !== 0 ? t : a.Agent.localeCompare(b.Agent);
  });

  // Build Con (TicketID-Agent) and dedup flags
  rows.forEach((r, i) => {
    r._con = `${r['Ticket Ticket ID']}-${r.Agent}`;
    r._dupAgent = i > 0 && rows[i - 1]._con === r._con;
    r._dupGrp = i > 0 && rows[i - 1]['Ticket Ticket ID'] === r['Ticket Ticket ID'];
  });

  // Agent config maps
  const agents = (perfConfig.agents || []).slice().sort((a, b) => (a.sort || 9999) - (b.sort || 9999));
  const l2Set = new Set(agents.filter(a => a.role === 'L2').map(a => a.name));
  const targetMap = Object.fromEntries(agents.map(a => [a.name, a.target || 0]));
  const sortMap = Object.fromEntries(agents.map(a => [a.name, a.sort || 9999]));

  const L1_GROUPS = ['ERP Helpdesk-L1'];
  const L2_GROUPS = ['ERP Helpdesk-L1', 'ERP Helpdesk-L2'];
  const DONE_STATUS = ['Resolved', 'Closed'];

  // ── Per-agent accumulators ────────────────────────────────────────────────
  const acc = {}; // { agentName: { resolved, forwarded, high, urgent, role } }
  const ensureAcc = (name, role) => {
    if (!acc[name]) acc[name] = { resolved: 0, forwarded: 0, high: 0, urgent: 0, role };
  };

  for (const r of rows) {
    const agent = r.Agent;
    if (!agent) continue;
    const role = l2Set.has(agent) ? 'L2' : 'L1';
    ensureAcc(agent, role);
    if (r._dupAgent) continue; // skip duplicate ticket-agent combos

    const grp = (r['Ticket Group name'] || '').trim();
    const status = (r['Ticket Status'] || '').trim();
    const pri = (r['Ticket Priority'] || '').trim().toLowerCase();
    const isResolved = DONE_STATUS.includes(status);

    if (role === 'L1') {
      const inL1Grp = L1_GROUPS.includes(grp);
      if (inL1Grp && isResolved) {
        acc[agent].resolved++;
        if (pri === 'high') acc[agent].high++;
        if (pri === 'urgent') acc[agent].urgent++;
      } else if (!inL1Grp) {
        acc[agent].forwarded++;
        if (pri === 'high') acc[agent].high++;
        if (pri === 'urgent') acc[agent].urgent++;
      }
    } else {
      // L2
      const inL2Grp = L2_GROUPS.includes(grp);
      if (inL2Grp && isResolved) {
        acc[agent].resolved++;
      } else if (!inL2Grp) {
        acc[agent].forwarded++;
      }
    }
  }

  // ── Team processed (unique ticket dedup by group) ─────────────────────────
  let teamProcessed = 0;
  for (const r of rows) {
    if (r._dupGrp) continue;
    const grp = (r['Ticket Group name'] || '').trim();
    const status = (r['Ticket Status'] || '').trim();
    if (L2_GROUPS.includes(grp) && DONE_STATUS.includes(status)) teamProcessed++;
    else if (!L2_GROUPS.includes(grp)) teamProcessed++;
  }

  // ── Determine report date from Clocked date ───────────────────────────────
  let reportDate = '';
  for (const r of rows) {
    const d = r['Clocked date'];
    if (d) { reportDate = d.substring(0, 10); break; }
  }

  // ── Build ordered agent rows ──────────────────────────────────────────────
  const allAgentNames = Object.keys(acc);
  allAgentNames.sort((a, b) => (sortMap[a] || 9999) - (sortMap[b] || 9999));

  const agentRows = allAgentNames.map(name => {
    const a = acc[name];
    const target = targetMap[name] || 0;
    const processed = a.resolved + a.forwarded;
    return {
      name,
      role: a.role,
      resolved: a.resolved,
      forwarded: a.forwarded,
      processed,
      high: a.role === 'L1' ? a.high : null,
      urgent: a.role === 'L1' ? a.urgent : null,
      target: target || null,
      hitTarget: target > 0 ? processed >= target : null,
      achievePct: target > 0 ? Math.round((processed / target) * 100) : null
    };
  });

  // ── L1 totals ─────────────────────────────────────────────────────────────
  const l1Rows = agentRows.filter(r => r.role === 'L1');
  const totals = {
    name: 'Total (L1)',
    role: 'TOTAL',
    resolved: l1Rows.reduce((s, r) => s + r.resolved, 0),
    forwarded: l1Rows.reduce((s, r) => s + r.forwarded, 0),
    processed: l1Rows.reduce((s, r) => s + r.processed, 0),
    high: l1Rows.reduce((s, r) => s + (r.high || 0), 0),
    urgent: l1Rows.reduce((s, r) => s + (r.urgent || 0), 0),
    target: l1Rows.reduce((s, r) => s + (r.target || 0), 0),
    hitTarget: null, achievePct: null
  };

  return {
    reportDate,
    fileName,
    agentRows,
    totals,
    teamProcessed,
    totalInward,
    closingUnresolved
  };
}

function splitCSVLine(line) {
  const result = []; let cur = ''; let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') { inQ = !inQ; }
    else if (c === ',' && !inQ) { result.push(cur); cur = ''; }
    else cur += c;
  }
  result.push(cur);
  return result;
}

// Health check endpoint (used by keep-alive ping and Render health checks)
app.get('/health', (_,r)=>r.json({status:'ok',ts:new Date().toISOString()}));

app.get('*', (_,r)=>r.sendFile(path.join(__dirname,'public','index.html')));
app.listen(PORT,()=>{
  console.log(`\n✅  C-Serv.AI  →  http://localhost:${PORT}`);
  console.log('    superadmin / Super@123 | admin / Admin@123 | operator / Op@123\n');

  // ── Keep-alive self-ping every 14 min (prevents Render free-tier sleep) ──
  const APP_URL = process.env.APP_URL || '';
  if(APP_URL){
    const https = require('https');
    const http  = require('http');
    setInterval(()=>{
      try{
        const mod = APP_URL.startsWith('https') ? https : http;
        mod.get(APP_URL + '/health', r => {
          r.resume();
          console.log(`[keep-alive] ping ${new Date().toISOString()} → ${r.statusCode}`);
        }).on('error', e => console.warn('[keep-alive] ping failed:', e.message));
      }catch(e){console.warn('[keep-alive] error:',e.message);}
    }, 14 * 60 * 1000); // every 14 minutes
    console.log(`[keep-alive] active → pinging ${APP_URL} every 14 min`);
  }
});
