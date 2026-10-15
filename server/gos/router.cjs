const express = require('express'),
  fs = require('node:fs'),
  path = require('node:path'),
  crypto = require('node:crypto');
const seed = require('./seed.json');
const catalogue = require('./catalogue/catalogue.json');
const {
  PDFDocument
} = require('pdf-lib');
const {
  matchDesign,
  designLabelKey,
  pagePdf
} = require('./pdf.cjs');
const clean = (v, n = 500) => String(v ?? '').trim().slice(0, n);
const regKey = v => clean(v, 20).toUpperCase().replace(/\s/g, '');
function imageBuffer(data) {
  const m = /^data:image\/(png|jpeg|webp);base64,([A-Za-z0-9+/=]+)$/.exec(data || '');
  if (!m) throw Error('Use a PNG, JPEG or WebP image.');
  const buffer = Buffer.from(m[2], 'base64');
  if (!buffer.length || buffer.length > 8 * 1024 * 1024) throw Error('Images must be under 8MB.');
  const type = m[1];
  if (type === 'png' && !buffer.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))) throw Error('Invalid PNG.');
  if (type === 'jpeg' && !(buffer[0] === 255 && buffer[1] === 216 && buffer[2] === 255)) throw Error('Invalid JPEG.');
  if (type === 'webp' && !(buffer.toString('ascii', 0, 4) === 'RIFF' && buffer.toString('ascii', 8, 12) === 'WEBP')) throw Error('Invalid WebP.');
  return {
    buffer,
    type,
    ext: type === 'jpeg' ? 'jpg' : type
  };
}
function createGosRouter({
  dataDir,
  sendMail,
  verifyMail,
  recipient = process.env.GOS_ORDERS_TO || "eddy.d'antonio@signs-express.co.uk",
  cc = process.env.GOS_ORDERS_CC || 'matt.rutlidge@signs-express.co.uk',
  password = process.env.GOS_CUSTOMER_PASSWORD || 'GOS123',
  secure = process.env.NODE_ENV === 'production',
  preview = false
} = {}) {
  const router = express.Router(),
    sessions = new Map(),
    attempts = new Map();
  let queue = Promise.resolve();
  const file = path.join(dataDir, 'register.json'),
    media = path.join(dataDir, 'photos');
  const documentsDir = path.join(dataDir, 'documents');
  fs.mkdirSync(documentsDir, {
    recursive: true
  });
  const masterDir = path.join(__dirname, 'catalogue');
  function docs(s) {
    const uploaded = s.documents || [];
    return [...(uploaded.some(d => d.id === catalogue.id) ? [] : [{ ...catalogue, bundled: true }]), ...uploaded];
  }
  function designs(s) {
    return docs(s).flatMap(doc => doc.pages.map(p => ({
      ...p,
      documentId: doc.id,
      id: p.id || doc.id + '-page-' + p.pageNumber,
      liveryId: p.liveryId || s.liveries.find(l => l.name === p.liveryName)?.id,
      previewUrl: '/api/gos/documents/' + doc.id + '/pages/' + p.pageNumber + '/preview',
      pdfUrl: '/api/gos/documents/' + doc.id + '/pages/' + p.pageNumber + '.pdf',
      documentUrl: '/api/gos/documents/' + doc.id + '.pdf'
    })));
  }
  function documentFile(doc) {
    return path.join(doc.bundled ? masterDir : documentsDir, doc.filename);
  }
  function findDesign(s, id) {
    return designs(s).find(d => d.id === id);
  }
  fs.mkdirSync(media, {
    recursive: true
  });
  function read() {
    if (!fs.existsSync(file)) write(structuredClone(seed));
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  }
  function write(store) {
    const temp = file + '.tmp';
    fs.writeFileSync(temp, JSON.stringify(store, null, 2));
    if (fs.existsSync(file)) fs.copyFileSync(file, file + '.previous');
    fs.renameSync(temp, file);
  }
  function mutate(fn) {
    const next = queue.then(async () => {
      const store = read();
      const result = await fn(store);
      write(store);
      return result;
    });
    queue = next.catch(() => {});
    return next;
  }
  const route = fn => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(e => res.status(e.status || 400).json({
    error: e.message || 'Request failed.'
  }));
  router.use((req, res, next) => {
    res.set('Cache-Control', 'no-store');
    res.set('X-Content-Type-Options', 'nosniff');
    if (req.method !== 'GET' && req.headers.origin) {
      try {
        if (new URL(req.headers.origin).host !== req.headers.host) return res.status(403).json({
          error: 'Request origin rejected.'
        });
      } catch {
        return res.status(403).json({
          error: 'Request origin rejected.'
        });
      }
    }
    next();
  });
  router.post('/login', (req, res) => {
    const key = req.ip,
      now = Date.now();
    let a = attempts.get(key);
    if (!a || a.until < now) {
      a = {
        count: 0,
        until: now + 900000
      };
      attempts.set(key, a);
    }
    if (++a.count > 10) return res.status(429).json({
      error: 'Too many attempts. Try again in 15 minutes.'
    });
    const given = crypto.createHash('sha256').update(String(req.body.password || '')).digest(),
      expected = crypto.createHash('sha256').update(password).digest();
    if (!crypto.timingSafeEqual(given, expected)) return res.status(401).json({
      error: 'Incorrect password.'
    });
    attempts.delete(key);
    const token = crypto.randomBytes(32).toString('hex');
    sessions.set(token, now + 12 * 3600000);
    res.cookie('gos_session', token, {
      httpOnly: true,
      secure,
      sameSite: 'strict',
      maxAge: 12 * 3600000,
      path: '/'
    });
    res.json({
      ok: true
    });
  });
  router.use((req, res, next) => {
    for (const [sessionToken, expires] of sessions) if (expires < Date.now()) sessions.delete(sessionToken);
    const token = (req.headers.cookie || '').split(';').map(s => s.trim()).find(s => s.startsWith('gos_session='))?.slice(12);
    if (!token || !(sessions.get(token) > Date.now())) return res.status(401).json({
      error: 'GOS password required.'
    });
    req.gosToken = token;
    next();
  });
  router.post('/logout', (req, res) => {
    sessions.delete(req.gosToken);
    res.clearCookie('gos_session', {
      path: '/'
    });
    res.json({
      ok: true
    });
  });
  const photoPublic = p => ({
    id: p.id,
    name: p.name,
    url: '/api/gos/photos/' + p.file
  });
  function view(s) {
    const { documentHistory, deletedVehicles, ...publicStore } = s;
    return {
      ...publicStore,
      vehicles: s.vehicles.map(v => ({
        ...v,
        matchedDesign: matchDesign(v, designs(s)),
        photos: (v.photos || []).map(photoPublic)
      })),
      liveries: s.liveries.map(l => ({
        ...l,
        photos: (l.photos || []).map(photoPublic)
      })),
      designs: designs(s),
      documents: docs(s).filter(d => !d.vehicleId).map(d => ({
        id: d.id,
        name: d.name,
        pageCount: d.pages.length,
        revision: d.filename,
        url: '/api/gos/documents/' + d.id + '.pdf'
      })),
      orders: s.orders.map(o => ({
        id: o.id,
        kind: o.kind,
        reg: o.vehicle.reg,
        livery: o.livery.name,
        createdAt: o.createdAt,
        status: o.status
      })),
      orderingReady: Boolean(recipient && sendMail),
      previewMode: preview
    };
  }
  router.get('/register', route((req, res) => res.json(view(read()))));
  let mailStatusCache = null,
    mailStatusPromise = null;
  router.get('/email-status', route(async (req, res) => {
    if (!verifyMail) return res.json({
      ready: Boolean(recipient && sendMail),
      verified: false,
      to: recipient,
      cc
    });
    if (mailStatusCache && mailStatusCache.until > Date.now()) return res.status(mailStatusCache.code).json(mailStatusCache.result);
    if (!mailStatusPromise) mailStatusPromise = (async () => {
      try {
        await verifyMail();
        return {
          until: Date.now() + 300000,
          code: 200,
          result: {
            ready: true,
            verified: true,
            to: recipient,
            cc
          }
        };
      } catch {
        return {
          until: Date.now() + 30000,
          code: 503,
          result: {
            ready: false,
            verified: false,
            error: 'The email connection could not be verified.'
          }
        };
      }
    })().finally(() => {
      mailStatusPromise = null;
    });
    mailStatusCache = await mailStatusPromise;
    res.status(mailStatusCache.code).json(mailStatusCache.result);
  }));
  router.get('/photos/:file', route((req, res) => {
    if (!/^[a-f0-9-]+\.(jpg|png|webp)$/.test(req.params.file)) return res.status(404).end();
    res.sendFile(path.join(media, req.params.file));
  }));
  function vehicleFields(body, s, old = {}) {
    const v = {
      ...old
    };
    for (const k of ['reg', 'name', 'division', 'year', 'make', 'model', 'group', 'wheelbase', 'colour', 'liveryId', 'designId', 'notes']) v[k] = clean(body[k] ?? old[k] ?? '', k === 'notes' ? 3000 : 300);
    v.reg = v.reg.toUpperCase();
    if (!v.reg || !v.model) throw Error('Registration and make/model are required.');
    if (v.designId && !designs(s).some(d => d.id === v.designId && d.liveryId === v.liveryId && (!d.vehicleId || d.vehicleId === v.id))) throw Error('Select a PDF design in this livery category.');
    if (!s.liveries.some(l => l.id === v.liveryId)) throw Error('Choose a livery.');
    if (s.vehicles.some(other => other.id !== old.id && regKey(other.reg) === regKey(v.reg))) throw Error('That registration already exists.');
    return v;
  }
  router.post('/vehicles', route(async (req, res) => {
    const v = await mutate(s => {
      const v = {
        ...vehicleFields(req.body, s),
        id: crypto.randomUUID(),
        photos: [],
        version: 1
      };
      s.vehicles.push(v);
      return v;
    });
    res.status(201).json({
      id: v.id
    });
  }));
  router.patch('/vehicles/:id', route(async (req, res) => {
    await mutate(s => {
      const i = s.vehicles.findIndex(v => v.id === req.params.id);
      if (i < 0) throw Error('Vehicle not found.');
      if (req.body.version !== s.vehicles[i].version) {
        const e = Error('This vehicle has changed. Reload before editing.');
        e.status = 409;
        throw e;
      }
      s.vehicles[i] = {
        ...vehicleFields(req.body, s, s.vehicles[i]),
        version: s.vehicles[i].version + 1
      };
    });
    res.json({
      ok: true
    });
  }));
  router.delete('/vehicles/:id', route(async (req, res) => {
    await mutate(s => {
      const vehicle = s.vehicles.find(v => v.id === req.params.id);
      if (!vehicle) throw Error('Van not found.');
      if (req.body.version !== vehicle.version) {
        const error = Error('This van has changed. Close and reopen it before deleting.');
        error.status = 409;
        throw error;
      }
      s.deletedVehicles = [...(s.deletedVehicles || []), { ...vehicle, deletedAt: new Date().toISOString() }];
      s.vehicles = s.vehicles.filter(v => v.id !== vehicle.id);
    });
    res.json({ ok: true });
  }));
  router.post('/liveries', route(async (req, res) => {
    const name = clean(req.body.name, 120);
    if (!name) throw Error('Livery name required.');
    const id = await mutate(s => {
      if (s.liveries.some(l => l.name.toLowerCase() === name.toLowerCase())) throw Error('Livery name already exists.');
      const id = crypto.randomUUID();
      s.liveries.push({
        id,
        name,
        photos: []
      });
      return id;
    });
    res.status(201).json({
      id
    });
  }));
  router.post('/:kind/:id/photos', route(async (req, res) => {
    if (!['vehicles', 'liveries'].includes(req.params.kind)) throw Error('Invalid photo destination.');
    const decoded = imageBuffer(req.body.dataUrl);
    await mutate(s => {
      const target = s[req.params.kind].find(v => v.id === req.params.id);
      if (!target) throw Error('Item not found.');
      if ((target.photos || []).length >= 8) throw Error('Maximum eight photos per item.');
      const id = crypto.randomUUID(),
        filename = id + '.' + decoded.ext;
      fs.writeFileSync(path.join(media, filename), decoded.buffer);
      target.photos = [...(target.photos || []), {
        id,
        file: filename,
        name: clean(req.body.name, 150) || 'Photo'
      }];
      if (req.params.kind === 'vehicles') target.version++;
    });
    res.json({
      ok: true
    });
  }));
  router.get('/documents/:docId/pages/:pageNumber/preview', route((req, res) => {
    const s = read(),
      doc = docs(s).find(d => d.id === req.params.docId),
      page = doc?.pages.find(p => p.pageNumber === Number(req.params.pageNumber));
    if (!page) return res.status(404).end();
    res.sendFile(path.join(doc.bundled ? masterDir : documentsDir, page.previewFile));
  }));
  router.get('/documents/:docId/pages/:page.pdf', route(async (req, res) => {
    const doc = docs(read()).find(d => d.id === req.params.docId);
    if (!doc) return res.status(404).end();
    const bytes = await pagePdf(fs.readFileSync(documentFile(doc)), Number(req.params.page));
    res.type('application/pdf').set('Content-Disposition', 'inline; filename="GOS-page-' + req.params.page + '.pdf"').send(bytes);
  }));
  router.get('/documents/:docId.pdf', route((req, res) => {
    const doc = docs(read()).find(d => d.id === req.params.docId);
    if (!doc) return res.status(404).end();
    res.type('application/pdf').sendFile(documentFile(doc));
  }));
  router.post('/documents', route(async (req, res) => {
    const m = /^data:application\/pdf;base64,([A-Za-z0-9+/=]+)$/.exec(req.body.dataUrl || '');
    if (!m) throw Error('Upload a PDF document.');
    const bytes = Buffer.from(m[1], 'base64');
    if (bytes.length > 22 * 1024 * 1024 || bytes.toString('ascii', 0, 5) !== '%PDF-') throw Error('Use a valid PDF under 22MB.');
    const pdf = await PDFDocument.load(bytes);
    if (pdf.getPageCount() > 100) throw Error('A PDF can have at most 100 pages.');
    if (!Array.isArray(req.body.pages) || req.body.pages.length !== pdf.getPageCount()) throw Error('Each PDF page needs a preview and van label.');
    const uploadId = crypto.randomUUID(), s = read();
    if (req.body.vehicleId && (req.body.replaceId || pdf.getPageCount() !== 1)) throw Error('Choose a one-page PDF for this van.');
    const previous = req.body.vehicleId ? docs(s).find(d => d.vehicleId === req.body.vehicleId) : req.body.replaceId ? docs(s).find(d => d.id === req.body.replaceId) : null;
    if (previous?.vehicleId && !req.body.vehicleId) throw Error('Replace this PDF from the van editor.');
    if (req.body.replaceId && !previous) throw Error('PDF not found. Refresh and try again.');
    const id = previous?.id || uploadId;
    const pages = req.body.pages.map((p, i) => {
      if (p.pageNumber !== i + 1 || !s.liveries.some(l => l.id === p.liveryId)) throw Error('Choose a livery category for every page.');
      const preview = imageBuffer(p.previewDataUrl),
        file = uploadId + '-page-' + (i + 1) + '.' + preview.ext;
      fs.writeFileSync(path.join(documentsDir, file), preview.buffer);
      return {
        id: uploadId + '-page-' + (i + 1),
        pageNumber: i + 1,
        title: clean(p.title, 500) || 'Page ' + (i + 1),
        liveryId: p.liveryId,
        previewFile: file,
        ...(req.body.vehicleId ? {vehicleId: clean(req.body.vehicleId, 80)} : {})
      };
    });
    fs.writeFileSync(path.join(documentsDir, uploadId + '.pdf'), bytes);
    await mutate(s => {
      const van = req.body.vehicleId && s.vehicles.find(v => v.id === req.body.vehicleId);
      if (req.body.vehicleId && (!van || van.version !== req.body.vehicleVersion)) {
        const error = Error('This van has changed. Close and reopen it before replacing the PDF.');
        error.status = 409; throw error;
      }
      const old = previous && docs(s).find(d => d.id === previous.id);
      if (previous && (!old || old.filename !== (van ? previous.filename : req.body.revision))) {
        const error = Error('This PDF has changed. Refresh before replacing it.');
        error.status = 409;
        throw error;
      }
      if (old) {
        const allDesigns = designs(s);
        const oldDesigns = allDesigns.filter(d => d.documentId === old.id);
        // Match labels, not positions, so reordered pages cannot change a van's artwork.
        const key = p => designLabelKey(p.title) + '|' + p.liveryId;
        for (const page of pages) {
          const matches = oldDesigns.filter(d => key(d) === key(page));
          if (matches.length === 1 && pages.filter(p => key(p) === key(page)).length === 1) page.id = matches[0].id;
        }
        for (const v of s.vehicles) {
          const assigned = matchDesign(v, allDesigns);
          if (assigned?.documentId === old.id && v.id !== van?.id) {
            // Removed drawings require a manual check instead of guessing another drawing.
            v.designId = assigned.id;
            v.version++;
          }
        }
        s.documentHistory = [...(s.documentHistory || []), { ...old, replacedAt: new Date().toISOString() }];
      }
      s.documents = [...(s.documents || []).filter(d => d.id !== id), {
        id, name: clean(req.body.name, 160) || 'Livery PDF', filename: uploadId + '.pdf', pages,
        ...(van ? {vehicleId: van.id} : {})
      }];
      if (van) {
        const fields = vehicleFields({...req.body.vehicle, designId: pages[0].id}, s, van);
        s.vehicles[s.vehicles.findIndex(v => v.id === van.id)] = {...fields, version: van.version + 1};
      }
    });
    res.json({
      id
    });
  }));
  router.post('/orders', route(async (req, res) => {
    if (!recipient || !sendMail) {
      const e = Error('The order inbox is not configured yet. Please contact Signs Express.');
      e.status = 503;
      throw e;
    }
    const body = req.body,
      requestId = clean(body.requestId, 80);
    if (!/^[a-zA-Z0-9-]{10,80}$/.test(requestId)) throw Error('Order reference missing.');
    let order = await mutate(s => {
      const existing = s.orders.find(o => o.requestId === requestId);
      if (existing) return existing;
      if (!['livery', 'repair'].includes(body.kind)) throw Error('Choose a livery order or repair order.');
      const contact = clean(body.contact, 160) || 'GOS customer portal',
        email = clean(body.email, 200);
      if (email && (!/^\S+@\S+\.\S+$/.test(email) || /[\r\n]/.test(email))) throw Error('Invalid reply email.');
      let vehicle;
      if (body.existingId) {
        vehicle = s.vehicles.find(v => v.id === body.existingId);
        if (!vehicle) throw Error('Vehicle not found.');
        if (body.vehicleVersion !== undefined && body.vehicleVersion !== vehicle.version) throw Error('This vehicle has changed. Close and reopen the order to refresh it.');
        vehicle = structuredClone(vehicle);
      } else {
        vehicle = {
          ...vehicleFields(body.vehicle || {}, s),
          id: crypto.randomUUID(),
          photos: [],
          version: 1
        };
        if (!vehicle.make || !vehicle.wheelbase || !vehicle.colour) throw Error('Make, model, wheelbase and colour are required for a new registration.');
      }
      const livery = s.liveries.find(l => l.id === (body.existingId ? vehicle.liveryId : body.liveryId));
      if (!livery) throw Error('Choose a livery.');
      const design = body.existingId ? matchDesign(vehicle, designs(s)) : body.designId ? findDesign(s, body.designId) : null;
      if (body.designId && (!design || design.liveryId !== livery.id || (design.vehicleId && design.vehicleId !== vehicle.id))) throw Error('Select a PDF design from the chosen livery.');
      if (design && !body.existingId) vehicle.designId = design.id;
      const photos = vehicle.liveryId === livery.id && vehicle.photos?.length ? vehicle.photos : livery.photos;
      if (!design && !photos?.length) throw Error('Upload a vehicle or livery photo before ordering so the email can include the image.');
      let annotations = [];
      if (body.kind === 'repair') {
        if (!Array.isArray(body.annotations) || !body.annotations.length || body.annotations.length > 8) throw Error('Mark at least one repair area.');
        annotations = body.annotations.map(a => {
          if (design && a.designId !== design.id) throw Error('Repair markings do not match the selected PDF page.');
          if (!design && !photos.some(p => p.id === a.photoId)) throw Error('Repair image is not linked to this vehicle or livery.');
          if (!Array.isArray(a.boxes) || !a.boxes.length || a.boxes.length > 50 || !a.boxes.every(b => ['x', 'y', 'w', 'h'].every(k => Number.isFinite(b[k]) && b[k] >= 0 && b[k] <= 1) && b.w > 0 && b.h > 0 && b.x + b.w <= 1.001 && b.y + b.h <= 1.001)) throw Error('Invalid repair boxes.');
          if (design) return {
            designId: design.id,
            pageNumber: design.pageNumber,
            boxes: a.boxes
          };
          const decoded = imageBuffer(a.dataUrl);
          const file = crypto.randomUUID() + '.' + decoded.ext;
          fs.writeFileSync(path.join(media, file), decoded.buffer);
          return {
            photoId: a.photoId,
            boxes: a.boxes,
            file
          };
        });
      }
      const order = {
        id: crypto.randomUUID(),
        requestId,
        kind: body.kind,
        vehicle,
        livery: structuredClone(livery),
        contact,
        email,
        notes: clean(body.notes, 5000),
        annotations,
        design: design ? {
          ...design,
          filename: docs(s).find(d => d.id === design.documentId).filename,
          bundled: docs(s).find(d => d.id === design.documentId).bundled
        } : null,
        createdAt: new Date().toISOString(),
        status: 'pending'
      };
      s.orders.push(order);
      return order;
    });
    if (order.status === 'sent') return res.json({
      ok: true,
      id: order.id,
      duplicate: true
    });
    if (order.status === 'sending') return res.status(409).json({
      error: 'This order is already being sent. Check the order list before retrying.'
    });
    await mutate(s => {
      const live = s.orders.find(o => o.id === order.id);
      if (live.status === 'sending' || live.status === 'sent') throw Error('Order already submitted.');
      live.status = 'sending';
    });
    try {
      const photos = order.vehicle.liveryId === order.livery.id && order.vehicle.photos?.length ? order.vehicle.photos : order.livery.photos;
      const files = order.design ? [] : order.kind === 'repair' ? order.annotations : photos;
      let attachments = files.map((p, i) => ({
        filename: (order.kind === 'repair' ? 'repair-marked-' : 'livery-') + (i + 1) + path.extname(p.file),
        path: path.join(media, p.file)
      }));
      if (order.design) {
        const boxes = order.kind === 'repair' ? order.annotations.flatMap(a => a.boxes) : [];
        const content = await pagePdf(fs.readFileSync(documentFile(order.design)), order.design.pageNumber, boxes);
        attachments = [{
          filename: (order.kind === 'repair' ? 'GOS-repair-marked-' : 'GOS-livery-') + order.vehicle.reg.replace(/[^a-z0-9]/gi, '') + '.pdf',
          content,
          contentType: 'application/pdf'
        }];
      }
      if (order.kind === 'repair') attachments.push({
        filename: 'repair-areas.json',
        content: JSON.stringify(order.annotations.map(a => ({
          photoId: a.photoId,
          boxes: a.boxes
        })), null, 2),
        contentType: 'application/json'
      });
      const details = [['Order', order.id], ['Type', order.kind === 'repair' ? 'Repair' : 'Vehicle livery'], ['Registration', order.vehicle.reg], ['Make', order.vehicle.make], ['Model', order.vehicle.model], ['Wheelbase / roof', order.vehicle.wheelbase], ['Colour', order.vehicle.colour], ['Division', order.vehicle.division], ['Driver', order.vehicle.name], ['Livery', order.livery.name], ['PDF design', order.design?.title], ['Source PDF page', order.design?.pageNumber], ['Requested via', 'GOS customer portal'], ['Notes', order.notes]].map(([k, v]) => k + ': ' + (v || 'Not supplied')).join('\n');
      const result = await sendMail({
        to: recipient,
        cc,
        replyTo: order.email || undefined,
        subject: 'GOS ' + (order.kind === 'repair' ? 'repair' : 'livery') + ' order — ' + order.vehicle.reg,
        text: details,
        attachments
      });
      if (result?.rejected?.length) throw Error('The mail server rejected delivery.');
      await mutate(s => {
        const saved = s.orders.find(o => o.id === order.id);
        saved.status = 'sent';
        saved.sentAt = new Date().toISOString();
        saved.messageId = result?.messageId || '';
        if (!s.vehicles.some(v => regKey(v.reg) === regKey(order.vehicle.reg)) && !(s.deletedVehicles || []).some(v => v.id === order.vehicle.id)) s.vehicles.push(order.vehicle);
      });
      res.json({
        ok: true,
        id: order.id
      });
    } catch (e) {
      await mutate(s => {
        const saved = s.orders.find(o => o.id === order.id);
        saved.status = 'failed';
        saved.error = clean(e.message, 500);
      });
      res.status(502).json({
        error: 'Email could not be confirmed as sent. Your request is saved as failed; retry or contact Signs Express.',
        id: order.id
      });
    }
  }));
  return router;
}
module.exports = {
  createGosRouter,
  imageBuffer,
  regKey
};
