# C-Serv.AI — Deployment Guide

## 🚀 Deploy to Render.com (FREE — 5 minutes)

> This gives you a real public URL like `https://cserv-ai.onrender.com`
> accessible from anywhere, on any device, 24/7.

---

## STEP 1 — Create a free GitHub account
Go to https://github.com and sign up (free). You'll need this to connect to Render.

---

## STEP 2 — Upload your project to GitHub

1. Go to https://github.com/new
2. Create a new repository named `cserv-ai`
3. Set it to **Private** (recommended)
4. Click **Create repository**
5. On the next screen, click **uploading an existing file**
6. Drag & drop ALL files from the `cserv-ai` folder:
   - `server.js`
   - `package.json`
   - `public/index.html`
   - `README.md`
7. Click **Commit changes**

---

## STEP 3 — Deploy on Render.com

1. Go to https://render.com and sign up (free, use GitHub login)
2. Click **New → Web Service**
3. Connect your GitHub account and select `cserv-ai` repository
4. Fill in settings:
   - **Name**: `cserv-ai` (or any name you like)
   - **Region**: Singapore (closest for India)
   - **Branch**: `main`
   - **Build Command**: `npm install`
   - **Start Command**: `npm start`
   - **Plan**: Free
5. Click **Create Web Service**
6. Wait ~2 minutes for deployment
7. Your app URL appears at the top: `https://cserv-ai.onrender.com`

---

## STEP 4 — First Login

Open your URL and login with:

| Role     | Username   | Password    |
|----------|------------|-------------|
| Admin    | `admin`    | `Admin@123` |
| Operator | `operator` | `Op@123`    |

⚠️ **Change these passwords immediately after first login!**

---

## What you get

### Admin Login
- Full dashboard access
- All module cards visible
- Create/delete/edit operator accounts
- Set which modules each operator can see
- Grant operators permission to Save and/or Export rosters
- Change passwords
- Edit agent list (saved permanently)

### Operator Login
- Only sees modules admin has enabled
- Cannot access Admin Panel
- Cannot edit agent list
- Can save/export only if admin granted permission

---

## Default Credentials Detail

```
Admin:    username=admin    password=Admin@123
Operator: username=operator password=Op@123
```

---

## Adding More Operators

1. Login as Admin
2. Go to **Admin Panel** (⚙️ in sidebar)
3. Click **+ Add User**
4. Set role to **Operator**
5. Share the username/password with your team member
6. Use **Operator Module Access** toggles to control what they see

---

## Data Storage

All data is stored in `data.json` on the server:
- Users and passwords (encrypted with bcrypt)
- Agent list
- Saved rosters
- Module access settings

On Render free tier, data persists as long as the service is running.
For permanent data persistence, upgrade to Render's paid plan or add a database.

---

## Custom Domain (Optional)

1. In Render dashboard → your service → Settings → Custom Domains
2. Add your domain e.g. `roster.yourcompany.com`
3. Point your DNS CNAME to the Render URL

---

## Notes

- Free tier on Render spins down after 15 min inactivity
  (first load after idle takes ~30 seconds to wake up)
- For always-on, upgrade to Render Starter ($7/month)
- The app works on all mobile browsers (iOS Safari, Android Chrome, etc.)
