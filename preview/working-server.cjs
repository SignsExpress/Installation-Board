// Local rebuild only: production app, separate files, loopback listener.
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const express = require('express');
const root = path.resolve(__dirname, '..');
const runtime = path.join(root, 'outputs', 'rebuild-runtime');
const source = process.env.PORTAL_SOURCE_DIR;
if (!source) throw new Error('PORTAL_SOURCE_DIR must identify the original checkout.');
if (path.resolve(source) === root) throw new Error('Source must be a separate checkout.');
fs.mkdirSync(runtime, {recursive:true});
const read = p => JSON.parse(fs.readFileSync(p, 'utf8'));
const write = (name,data) => fs.writeFileSync(path.join(runtime,name),JSON.stringify(data,null,2));
if (!fs.existsSync(path.join(runtime,'jobs.json'))) {
  const backup = process.env.PORTAL_PREVIEW_BACKUP ? read(process.env.PORTAL_PREVIEW_BACKUP) : null;
  const captured = !backup && process.env.PORTAL_PREVIEW_SNAPSHOT ? read(process.env.PORTAL_PREVIEW_SNAPSHOT) : null;
  const original = backup?.data?.boardStore || read(path.join(source,'data','jobs.json'));
  const board = structuredClone(original);
  if (captured) {
    board.jobs = captured.board.jobs.map(j=>({...j,id:crypto.randomUUID(),notes:j.notes || ''}));
    board.designBoard = {...board.designBoard,cards:captured.board.designBoard.cards.map(j=>{
      const [day,month,year] = j.date.split('/');
      return {...j,id:crypto.randomUUID(),status:({'New Orders':'new','Awaiting Sign-Off':'awaiting-sign-off','Order with Salesperson':'order-with-salesperson'})[j.status] || 'new',createdAt:year+'-'+month+'-'+day+'T09:00:00Z',designerNote:j.notes,jobTotalExVat:Number(j.displayValue.replace(/[£,]/g,''))};
    })};
  }
  // Never deliver queued production messages from an imported backup.
  board.socialPostQueue = [];
  board.pushSubscriptions = [];
  write('jobs.json',board);
  write('installers.json',backup?.data?.installers || read(path.join(source,'data','installers-live.json')));
  write('requests.json',backup?.data?.requests || []);
  const users = structuredClone(backup?.data?.usersStore || read(path.join(source,'data','users.json')));
  write('users.json',users);
  const sourceAssets = path.join(source,'data','pro-forma-assets');
  if (fs.existsSync(sourceAssets)) fs.cpSync(sourceAssets,path.join(runtime,'pro-forma-assets'),{recursive:true});
}
process.env.DATA_FILE=path.join(runtime,'jobs.json');
process.env.AUTH_USERS_FILE=path.join(runtime,'users.json');
process.env.INSTALLERS_FILE=path.join(runtime,'installers.json');
process.env.REQUESTS_FILE=path.join(runtime,'requests.json');
process.env.PORT='5186';
process.env.HOST='127.0.0.1';
process.env.PUBLIC_APP_ORIGIN='http://127.0.0.1:5186';
process.env.COREBRIDGE_SYNC_ENABLED='false';
delete process.env.AUTH_BOOTSTRAP_PASSWORDS;
for (const key of Object.keys(process.env)) if (/^(SMTP_|PUSH_VAPID_|TIMEMOTO_)/.test(key)) delete process.env[key];
const {hashPassword}=require('../server/auth-store');
const accessFile=path.join(runtime,'preview-access.json');
if (!fs.existsSync(accessFile)) {
  const users=read(process.env.AUTH_USERS_FILE);
  const owner=users.users.find(u=>u.displayName==='Matt Rutlidge') || users.users[0];
  if (!owner) throw new Error('No user profile available for the test copy.');
  const password=crypto.randomBytes(18).toString('base64url');
  Object.assign(owner,hashPassword(password));
  write('users.json',users);
  write('preview-access.json',{displayName:owner.displayName,password});
}
const app=express();
app.use((req,res,next)=>{
  if (/^\/(?:client\/)?(?:filtering|materials|mustang|morning-meeting)(?:\/|$)/.test(req.path)) return res.redirect('/');
  if (req.path.startsWith('/api/') && req.method!=='GET' && /(?:email|send|push|subscribe|queue)/i.test(req.path)) return res.status(409).json({error:'Sending and publishing are disconnected in the rebuild test workspace.'});
  next();
});
app.use(require('../server/index').createServer());
app.listen(5186,'127.0.0.1',()=>console.log('Working rebuild: http://127.0.0.1:5186 · Separate test data'));

