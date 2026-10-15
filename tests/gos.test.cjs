const test = require('node:test'),
  assert = require('node:assert/strict'),
  express = require('express'),
  fs = require('node:fs'),
  os = require('node:os'),
  path = require('node:path');
const {
  createGosRouter
} = require('../server/gos/router.cjs');
const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a/NkAAAAASUVORK5CYII=';
async function fixture(fn, options = {}) {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'gos-test-')),
    mails = [];
  const app = express();
  app.use(express.json({
    limit: '15mb'
  }));
  app.use('/api/gos', createGosRouter({
    dataDir,
    secure: false,
    recipient: 'test@example.invalid',
    sendMail: async message => {
      mails.push(message);
      return {
        messageId: 'test-id',
        rejected: []
      };
    },
    ...options
  }));
  app.get('/api/jobs', (req, res) => res.status(401).json({
    error: 'Portal login required.'
  }));
  const server = app.listen(0, '127.0.0.1');
  await new Promise(r => server.once('listening', r));
  const url = 'http://127.0.0.1:' + server.address().port;
  async function req(route, body, cookie, method = body ? 'POST' : 'GET') {
    const r = await fetch(url + route, {
      method,
      headers: {
        ...(cookie ? {
          cookie
        } : {}),
        ...(body ? {
          'Content-Type': 'application/json'
        } : {})
      },
      body: body ? JSON.stringify(body) : undefined
    });
    return {
      status: r.status,
      body: r.headers.get('content-type')?.includes('application/json') ? await r.json() : Buffer.from(await r.arrayBuffer()),
      contentType:r.headers.get('content-type'),
      cookie: r.headers.get('set-cookie')?.split(';')[0]
    };
  }
  try {
    await fn({
      req,
      dataDir,
      mails
    });
  } finally {
    await new Promise(r => server.close(r));
    fs.rmSync(dataDir, {
      recursive: true,
      force: true
    });
  }
}
test('GOS cookie grants fleet access only; spreadsheet seed and edit conflicts are preserved', () => fixture(async ({
  req
}) => {
  assert.equal((await req('/api/gos/register')).status, 401);
  assert.equal((await req('/api/gos/login', {
    password: 'wrong'
  })).status, 401);
  const login = await req('/api/gos/login', {
    password: 'GOS123'
  });
  assert.equal(login.status, 200);
  assert.equal((await req('/api/jobs', null, login.cookie)).status, 401);
  const r = await req('/api/gos/register', null, login.cookie);
  assert.equal(r.body.vehicles.length, 58);
  assert.equal(r.body.liveries.length, 5);
  const v = r.body.vehicles[0];
  assert.equal((await req('/api/gos/vehicles/' + v.id, {
    ...v,
    colour: 'White'
  }, login.cookie, 'PATCH')).status, 200);
  assert.equal((await req('/api/gos/vehicles/' + v.id, {
    ...v,
    colour: 'Red'
  }, login.cookie, 'PATCH')).status, 409);
  assert.equal((await req('/api/gos/vehicles', {
    ...v,
    id: undefined
  }, login.cookie)).status, 400);
  await req('/api/gos/logout', {}, login.cookie);
  assert.equal((await req('/api/gos/register', null, login.cookie)).status, 401);
}));
test('orders attach reference images; repair orders attach drawings; retries do not duplicate mail', () => fixture(async ({
  req,
  mails
}) => {
  const c = (await req('/api/gos/login', {
    password: 'GOS123'
  })).cookie;
  let r = (await req('/api/gos/register', null, c)).body;
  const v = r.vehicles[0];
  const request = {
    requestId: 'gos-order-test-0001',
    kind: 'livery',
    existingId: v.id,
    liveryId: v.liveryId,
    contact: 'Test customer',
    email: 'customer@example.invalid',
    notes: 'Test only'
  };
  assert.equal((await req('/api/gos/orders', request, c)).status, 400);
  assert.equal((await req('/api/gos/vehicles/' + v.id + '/photos', {
    name: 'Vehicle.png',
    dataUrl: png
  }, c)).status, 200);
  r = (await req('/api/gos/register', null, c)).body;
  const photo = r.vehicles.find(x => x.id === v.id).photos[0];
  assert.equal((await req('/api/gos/orders', request, c)).status, 200);
  assert.equal(mails.length, 1);
  assert.ok(fs.existsSync(mails[0].attachments[0].path));
  assert.match(mails[0].text, new RegExp(v.reg));
  assert.equal((await req('/api/gos/orders', request, c)).status, 200);
  assert.equal(mails.length, 1);
  const repair = {
    ...request,
    requestId: 'gos-repair-test-0002',
    kind: 'repair',
    annotations: [{
      photoId: photo.id,
      boxes: [{
        x: .1,
        y: .2,
        w: .3,
        h: .4
      }],
      dataUrl: png
    }]
  };
  assert.equal((await req('/api/gos/orders', repair, c)).status, 200);
  assert.equal(mails.length, 2);
  assert.match(mails[1].attachments[0].filename, /repair-marked/);
  assert.equal(mails[1].attachments[1].filename, 'repair-areas.json');
  assert.equal((await req('/api/gos/orders', {
    ...repair,
    requestId: 'gos-repair-invalid',
    annotations: [{
      ...repair.annotations[0],
      boxes: [{
        x: 1,
        y: .2,
        w: .3,
        h: .4
      }]
    }]
  }, c)).status, 400);
}));
test('SMTP failures are recorded as failed, not presented as successful', () => fixture(async ({
  req
}) => {
  const c = (await req('/api/gos/login', {
    password: 'GOS123'
  })).cookie;
  const r = (await req('/api/gos/register', null, c)).body,
    v = r.vehicles[0];
  await req('/api/gos/vehicles/' + v.id + '/photos', {
    dataUrl: png
  }, c);
  const response = await req('/api/gos/orders', {
    requestId: 'gos-failed-mail-0003',
    kind: 'livery',
    existingId: v.id,
    liveryId: v.liveryId,
    contact: 'Test',
    email: 'test@example.invalid'
  }, c);
  assert.equal(response.status, 502);
  const after = (await req('/api/gos/register', null, c)).body;
  assert.equal(after.orders[0].status, 'failed');
}, {
  sendMail: async () => {
    throw Error('SMTP offline');
  }
}));
test('a new-registration order validates vehicle details and adds the vehicle only after mail acceptance', () => fixture(async ({req,mails}) => {
 const cookie=(await req('/api/gos/login',{password:'GOS123'})).cookie;
 const register=(await req('/api/gos/register',null,cookie)).body;
 const livery=register.liveries[0];
 await req('/api/gos/liveries/'+livery.id+'/photos',{dataUrl:png,name:'Approved design'},cookie);
 const request={requestId:'gos-new-vehicle-0004',kind:'livery',liveryId:livery.id,vehicle:{reg:'GOS TEST',make:'Ford',model:'Transit Custom',wheelbase:'L2 H1',colour:'White',liveryId:livery.id},contact:'Test contact',email:'contact@example.invalid'};
 assert.equal((await req('/api/gos/orders',{...request,vehicle:{...request.vehicle,colour:''}},cookie)).status,400);
 assert.equal((await req('/api/gos/orders',request,cookie)).status,200);
 const after=(await req('/api/gos/register',null,cookie)).body;
 assert.equal(after.vehicles.length,59);
 assert.equal(after.vehicles.find(v=>v.reg==='GOS TEST').colour,'White');
 assert.equal(mails.length,1);
 assert.match(mails[0].text,/Wheelbase \/ roof: L2 H1/);
 assert.equal(mails[0].attachments.length,1);
}));
test('PDF source pages require GOS access, imports retain PDFs and repair email attachments are PDF',()=>fixture(async({req,mails})=>{
 const {PDFDocument}=require('pdf-lib'),{pagePdf}=require('../server/gos/pdf.cjs');
 assert.equal((await req('/api/gos/documents/gos-master-v1.pdf')).status,401);
 const cookie=(await req('/api/gos/login',{password:'GOS123'})).cookie;
 const register=(await req('/api/gos/register',null,cookie)).body;
 assert.equal(register.designs.length,19);
 const page=await req('/api/gos/documents/gos-master-v1/pages/1.pdf',null,cookie);
 assert.equal(page.status,200);assert.match(page.contentType,/application\/pdf/);assert.equal(page.body.subarray(0,5).toString(),'%PDF-');
 const uploaded=await req('/api/gos/documents',{name:'Test original.pdf',dataUrl:'data:application/pdf;base64,'+page.body.toString('base64'),pages:[{pageNumber:1,title:'Test model',liveryId:register.liveries[0].id,previewDataUrl:png}]},cookie);
 assert.equal(uploaded.status,200);assert.equal((await req('/api/gos/register',null,cookie)).body.designs.length,20);
 const vehicle=register.vehicles.find(v=>v.reg==='HW70 KVT'),design=register.designs.find(d=>d.pageNumber===1);
 const response=await req('/api/gos/orders',{requestId:'gos-pdf-repair-0005',kind:'repair',existingId:vehicle.id,liveryId:vehicle.liveryId,designId:design.id,annotations:[{designId:design.id,boxes:[{x:.4,y:.4,w:.2,h:.15}]}],contact:'PDF QA',email:'pdf@example.invalid'},cookie);
 assert.equal(response.status,200);assert.equal(mails.length,1);assert.equal(mails[0].attachments[0].contentType,'application/pdf');assert.match(mails[0].attachments[0].filename,/\.pdf$/);assert.equal((await PDFDocument.load(mails[0].attachments[0].content)).getPageCount(),1);
}));
