const express    = require('express');
const session    = require('express-session');
const bodyParser = require('body-parser');
const bcrypt     = require('bcryptjs');
const path       = require('path');
const fs         = require('fs');

const app  = express();
const PORT = process.env.PORT || 3000;
const DB   = path.join(__dirname, 'data.json');

function load() {
  if (!fs.existsSync(DB)) return initDB();
  try { return JSON.parse(fs.readFileSync(DB, 'utf8')); }
  catch { return initDB(); }
}
function save(d) { fs.writeFileSync(DB, JSON.stringify(d, null, 2)); }

function initDB() {
  const d = {
    users: [
      { id:1, username:'admin',    password:bcrypt.hashSync('Admin@123',10), role:'admin',    name:'Administrator', createdAt:new Date().toISOString() },
      { id:2, username:'operator', password:bcrypt.hashSync('Op@123',10),    role:'operator', name:'Team Operator',  allowedModules:['roster'], canSaveRoster:true, canExportRoster:true, createdAt:new Date().toISOString() }
    ],
    agents: [
      {emp:'EP0200',name:'Shrikant Nayak',      level:0,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0269',name:'Ritu Singh',           level:0,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0505',name:'Rohit Kumar Agarwal',  level:0,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0563',name:'Himanshi Khowal',      level:2,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0564',name:'Chetan Goel',          level:1,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0523',name:'Sushant Kumar Suman',  level:2,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0560',name:'Mohit Singh',          level:3,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0678',name:'Abhay Pratap',         level:4,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0726',name:'Swagata Bhoumik',      level:6,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0848',name:'Deepak Gupta',         level:7,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0442',name:'Shivam Garg',          level:1,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0524',name:'Anurag Tiwari',        level:4,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0567',name:'Triloki Varshney',     level:1,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0557',name:'Sujit Kumar',          level:3,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0741',name:'Amarnath Vishwakarma', level:5,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0673',name:'Dhruv Mishra',         level:5,dept:'Customer Service',loc:'Gurgaon-US'},
      {emp:'EP0798',name:'Naveen Kumar S',       level:6,dept:'Customer Service',loc:'Gurgaon-US'}
    ],
    rosters: [],
    moduleAccess: { roster:true },
    settings: { appName:'C-Serv.AI', orgName:'Customer Service Team' }
  };
  save(d); return d;
}

app.use(bodyParser.json({ limit:'5mb' }));
app.use(bodyParser.urlencoded({ extended:true }));
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({
  secret: process.env.SESSION_SECRET || 'cservai-2026-secret-key',
  resave:false, saveUninitialized:false,
  cookie:{ maxAge:8*60*60*1000, httpOnly:true }
}));

const auth  = (q,r,n) => q.session?.user ? n() : r.status(401).json({error:'Login required'});
const admin = (q,r,n) => q.session?.user?.role==='admin' ? n() : r.status(403).json({error:'Admin only'});

app.post('/api/login', (req,res) => {
  const {username,password} = req.body;
  const d = load();
  const u = d.users.find(x=>x.username===username);
  if (!u||!bcrypt.compareSync(password,u.password)) return res.status(401).json({error:'Invalid username or password'});
  req.session.user = { id:u.id, username:u.username, role:u.role, name:u.name,
    allowedModules:u.allowedModules||[], canSaveRoster:u.canSaveRoster||false, canExportRoster:u.canExportRoster||false };
  res.json({ ok:true, user:req.session.user, moduleAccess:d.moduleAccess });
});
app.post('/api/logout', (q,r) => { q.session.destroy(); r.json({ok:true}); });
app.get('/api/me', auth, (req,res) => {
  const d=load(); res.json({user:req.session.user, moduleAccess:d.moduleAccess, settings:d.settings});
});
app.get('/api/agents', auth, (_,r) => r.json(load().agents));
app.put('/api/agents', admin, (req,res) => { const d=load(); d.agents=req.body; save(d); res.json({ok:true}); });

app.get('/api/rosters', auth, (_,r) => {
  const d=load();
  r.json(d.rosters.map(x=>({id:x.id,title:x.title,month:x.month,year:x.year,agentCount:x.agentCount,targetWO:x.targetWO,savedAt:x.savedAt,savedBy:x.savedBy})));
});
app.post('/api/rosters', auth, (req,res) => {
  const u=req.session.user;
  if (u.role!=='admin'&&!u.canSaveRoster) return res.status(403).json({error:'Save permission denied'});
  const d=load();
  const r={...req.body, id:Date.now(), savedAt:new Date().toISOString(), savedBy:u.name};
  d.rosters.unshift(r); if(d.rosters.length>100) d.rosters=d.rosters.slice(0,100);
  save(d); res.json({ok:true,id:r.id});
});
app.get('/api/rosters/:id', auth, (req,res) => {
  const r=load().rosters.find(x=>x.id==req.params.id);
  if(!r) return res.status(404).json({error:'Not found'}); res.json(r);
});
app.delete('/api/rosters/:id', admin, (req,res) => {
  const d=load(); d.rosters=d.rosters.filter(x=>x.id!=req.params.id); save(d); res.json({ok:true});
});

app.get('/api/admin/users', admin, (_,r) => {
  r.json(load().users.map(u=>({id:u.id,username:u.username,role:u.role,name:u.name,allowedModules:u.allowedModules||[],createdAt:u.createdAt,canSaveRoster:u.canSaveRoster||false,canExportRoster:u.canExportRoster||false})));
});
app.post('/api/admin/users', admin, (req,res) => {
  const d=load(); const {username,password,name,role}=req.body;
  if(!username||!password||!name) return res.status(400).json({error:'username/password/name required'});
  if(d.users.find(u=>u.username===username)) return res.status(400).json({error:'Username already taken'});
  const u={id:Date.now(),username,name,password:bcrypt.hashSync(password,10),role:role||'operator',allowedModules:[],canSaveRoster:false,canExportRoster:false,createdAt:new Date().toISOString()};
  d.users.push(u); save(d); res.json({ok:true,id:u.id});
});
app.put('/api/admin/users/:id', admin, (req,res) => {
  const d=load(); const i=d.users.findIndex(u=>u.id==req.params.id);
  if(i<0) return res.status(404).json({error:'User not found'});
  const {name,password,role,allowedModules,canSaveRoster,canExportRoster}=req.body;
  if(name) d.users[i].name=name; if(role) d.users[i].role=role;
  if(allowedModules!==undefined) d.users[i].allowedModules=allowedModules;
  if(canSaveRoster!==undefined) d.users[i].canSaveRoster=canSaveRoster;
  if(canExportRoster!==undefined) d.users[i].canExportRoster=canExportRoster;
  if(password) d.users[i].password=bcrypt.hashSync(password,10);
  save(d); res.json({ok:true});
});
app.delete('/api/admin/users/:id', admin, (req,res) => {
  const d=load(); const u=d.users.find(x=>x.id==req.params.id);
  if(u?.username==='admin') return res.status(400).json({error:'Cannot delete default admin'});
  d.users=d.users.filter(x=>x.id!=req.params.id); save(d); res.json({ok:true});
});

app.get('/api/admin/access', admin, (_,r) => r.json(load().moduleAccess));
app.put('/api/admin/access', admin, (req,res) => { const d=load(); d.moduleAccess=req.body; save(d); res.json({ok:true}); });
app.put('/api/admin/settings', admin, (req,res) => { const d=load(); d.settings={...d.settings,...req.body}; save(d); res.json({ok:true}); });

app.get('*', (_,r) => r.sendFile(path.join(__dirname,'public','index.html')));

app.listen(PORT, () => {
  console.log(`\n✅  C-Serv.AI  →  http://localhost:${PORT}`);
  console.log('    Admin:    admin    /  Admin@123');
  console.log('    Operator: operator /  Op@123\n');
});
