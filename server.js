'use strict';
const express    = require('express');
const session    = require('express-session');
const bodyParser = require('body-parser');
const bcrypt     = require('bcryptjs');
const path       = require('path');
const fs         = require('fs');

const app  = express();
const PORT = process.env.PORT || 3000;
const DATA = path.join(__dirname, 'data.json');

/* ── helpers ── */
function load() {
  if (!fs.existsSync(DATA)) {
    const seed = {
      users: [
        { id:1, username:'admin',    password: bcrypt.hashSync('Admin@123',10),
          role:'admin',    name:'Administrator', createdAt: new Date().toISOString() },
        { id:2, username:'operator', password: bcrypt.hashSync('Op@123',10),
          role:'operator', name:'Team Operator', createdAt: new Date().toISOString(),
          allowedModules:['roster'] }
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
        {emp:'EP0798',name:'Naveen Kumar S',       level:6,dept:'Customer Service',loc:'Gurgaon-US'},
      ],
      rosters: [],
      moduleAccess: { roster:true },
      settings: { appName:'C-Serv.AI', orgName:'Customer Service Team' }
    };
    fs.writeFileSync(DATA, JSON.stringify(seed, null,2));
    return seed;
  }
  return JSON.parse(fs.readFileSync(DATA,'utf8'));
}
function save(db){ fs.writeFileSync(DATA, JSON.stringify(db,null,2)); }

/* ── middleware ── */
app.use(bodyParser.json({limit:'5mb'}));
app.use(bodyParser.urlencoded({extended:true}));
app.use(express.static(path.join(__dirname,'public')));
app.use(session({
  secret: process.env.SESSION_SECRET || 'cservai-2026',
  resave:false, saveUninitialized:false,
  cookie:{ maxAge: 8*60*60*1000, httpOnly:true,
           secure: process.env.NODE_ENV==='production' && process.env.FORCE_HTTPS==='true' }
}));

const auth  = (q,r,n)=>{ if(q.session?.user) return n(); r.status(401).json({error:'Login required'}); };
const admin = (q,r,n)=>{ if(q.session?.user?.role==='admin') return n(); r.status(403).json({error:'Admin only'}); };

/* ── AUTH ── */
app.post('/api/login',(req,res)=>{
  const db=load(), u=db.users.find(x=>x.username===req.body.username);
  if(!u||!bcrypt.compareSync(req.body.password||'',u.password))
    return res.status(401).json({error:'Invalid username or password'});
  req.session.user={id:u.id,username:u.username,role:u.role,name:u.name,allowedModules:u.allowedModules||[]};
  res.json({ok:true, user:req.session.user});
});
app.post('/api/logout',(req,res)=>{ req.session.destroy(()=>res.json({ok:true})); });
app.get('/api/me',auth,(req,res)=>{
  const db=load();
  res.json({user:req.session.user, moduleAccess:db.moduleAccess, settings:db.settings});
});

/* ── AGENTS ── */
app.get('/api/agents',  auth,  (req,res)=>res.json(load().agents));
app.put('/api/agents',  admin, (req,res)=>{ const db=load(); db.agents=req.body; save(db); res.json({ok:true}); });

/* ── ROSTERS ── */
app.get('/api/rosters', auth, (req,res)=>{
  const db=load();
  res.json(db.rosters.map(r=>({id:r.id,title:r.title,month:r.month,year:r.year,
    agentCount:r.agentCount,targetWO:r.targetWO,savedAt:r.savedAt,savedBy:r.savedBy})));
});
app.post('/api/rosters', auth, (req,res)=>{
  const db=load();
  const roster={...req.body, id:Date.now(),
    savedAt:new Date().toISOString(), savedBy:req.session.user.name};
  db.rosters.unshift(roster);
  if(db.rosters.length>30) db.rosters=db.rosters.slice(0,30);
  save(db); res.json({ok:true,id:roster.id});
});
app.get('/api/rosters/:id', auth, (req,res)=>{
  const r=load().rosters.find(x=>x.id==req.params.id);
  r ? res.json(r) : res.status(404).json({error:'Not found'});
});
app.delete('/api/rosters/:id', admin, (req,res)=>{
  const db=load(); db.rosters=db.rosters.filter(x=>x.id!=req.params.id);
  save(db); res.json({ok:true});
});

/* ── ADMIN: USERS ── */
app.get('/api/admin/users', admin, (req,res)=>{
  res.json(load().users.map(u=>({id:u.id,username:u.username,
    role:u.role,name:u.name,createdAt:u.createdAt,allowedModules:u.allowedModules||[]})));
});
app.post('/api/admin/users', admin, (req,res)=>{
  const db=load(); const {username,password,role,name}=req.body;
  if(!username||!password||!name) return res.status(400).json({error:'username, password, name required'});
  if(db.users.find(u=>u.username===username)) return res.status(400).json({error:'Username already exists'});
  const u={id:Date.now(),username,password:bcrypt.hashSync(password,10),
    role:role||'operator',name,allowedModules:[],createdAt:new Date().toISOString()};
  db.users.push(u); save(db); res.json({ok:true,id:u.id});
});
app.put('/api/admin/users/:id', admin, (req,res)=>{
  const db=load(); const i=db.users.findIndex(u=>u.id==req.params.id);
  if(i<0) return res.status(404).json({error:'User not found'});
  const {name,password,role,allowedModules}=req.body;
  if(name) db.users[i].name=name;
  if(role) db.users[i].role=role;
  if(allowedModules!==undefined) db.users[i].allowedModules=allowedModules;
  if(password) db.users[i].password=bcrypt.hashSync(password,10);
  save(db); res.json({ok:true});
});
app.delete('/api/admin/users/:id', admin, (req,res)=>{
  const db=load();
  if(db.users.find(u=>u.id==req.params.id&&u.username==='admin'))
    return res.status(400).json({error:'Cannot delete default admin'});
  db.users=db.users.filter(u=>u.id!=req.params.id); save(db); res.json({ok:true});
});

/* ── ADMIN: MODULE ACCESS ── */
app.get('/api/admin/access', admin, (req,res)=>res.json(load().moduleAccess));
app.put('/api/admin/access', admin, (req,res)=>{
  const db=load(); db.moduleAccess=req.body; save(db); res.json({ok:true});
});

/* ── ADMIN: SETTINGS ── */
app.put('/api/admin/settings', admin, (req,res)=>{
  const db=load(); db.settings={...db.settings,...req.body}; save(db); res.json({ok:true});
});

/* ── SPA fallback ── */
app.get('*',(req,res)=>res.sendFile(path.join(__dirname,'public','index.html')));

app.listen(PORT,()=>{
  console.log(`\n✅  C-Serv.AI  →  http://localhost:${PORT}`);
  console.log('    Admin   : admin    / Admin@123');
  console.log('    Operator: operator / Op@123\n');
});
