# SX Installation Board

Shared installation board built in the same hosted style as the `SX Installer Directory` app:

- React frontend
- Express backend
- JSON data storage
- real-time board refresh using Server-Sent Events
- rolling board view: previous week, current week, next week, and the week after
- UK bank holidays added automatically

## What it does

The board mirrors your paper pad layout:

- `Date` column on the left
- `Holidays` column in the middle
- `Jobs` column on the right
- weekday rows only, grouped into weeks
- shared live updates when someone adds, edits, or deletes a job

Each job includes:

- date
- title
- crew / initials
- category
- notes

## Local development

Install dependencies:

```powershell
npm install
```

Run the frontend and backend together:

```powershell
npm run dev
```

Open:

```text
http://localhost:5173
```

## Production on Render

This app is ready for a single Render web service.

Recommended settings:

- Build command: `npm install && npm run build`
- Start command: `npm start`
- Environment: `Node`
- Root directory: leave blank if this repo only contains this app

Important:

- live job data is stored in `data/jobs.json`
- on Render, add a persistent disk and point `DATA_FILE` to that disk path so jobs are not lost on redeploy

Example:

- Disk mount path: `/var/data`
- Environment variable: `DATA_FILE=/var/data/jobs.json`

Render docs:

- [Web Services](https://render.com/docs/web-services)
- [Deploying on Render](https://render.com/docs/deploys/)
- [Persistent Disks](https://render.com/docs/disks)

## Step-by-step GitHub + Render help

### 1. Create the GitHub repo

In this folder:

```powershell
git init
git add .
git commit -m "Initial installation board app"
```

Create a new empty GitHub repo, then connect and push:

```powershell
git remote add origin <your-github-repo-url>
git branch -M main
git push -u origin main
```

### 2. Create the Render service

1. Log in to Render
2. Click `New +`
3. Choose `Web Service`
4. Connect the GitHub repo
5. Render should detect Node automatically
6. Set:
   - Build command: `npm install && npm run build`
   - Start command: `npm start`
7. Add environment variable:
   - `DATA_FILE=/var/data/jobs.json`
8. Add a persistent disk mounted at:
   - `/var/data`
9. Deploy

### 3. Use the board

1. Open the hosted URL from Render
2. Add a job using the form on the left
3. The right-hand board updates immediately
4. Keep the app open on office screens, laptops, or tablets
5. Everyone sees the same live board

## Main files

- `server/index.js` - backend API, live events, bank holiday generation
- `src/App.jsx` - board UI and job editor
- `src/index.css` - board layout and styling
- `data/jobs.json` - saved jobs
