const test = require('node:test'),
  assert = require('node:assert/strict'),
  fs = require('node:fs'),
  path = require('node:path');
const {
  PDFDocument,
  degrees
} = require('pdf-lib');
const {
  matchDesign,
  pagePdf,
  pdfBox
} = require('../server/gos/pdf.cjs');
const catalogue = require('../server/gos/catalogue/catalogue.json'),
  seed = require('../server/gos/seed.json');
const designs = catalogue.pages.map(p => ({
  ...p,
  liveryId: seed.liveries.find(l => l.name === p.liveryName).id
}));
test('all supplied PDF pages use the requested category ranges and corner labels', () => {
  assert.equal(designs.length, 19);
  for (const d of designs) {
    assert.ok(d.title.includes('FORD') || d.title.includes('Ford'));
    assert.equal(d.liveryName, d.pageNumber <= 9 ? 'Standard GOS' : d.pageNumber <= 14 ? 'Mechanical Electrical' : d.pageNumber <= 16 ? 'Electrical' : 'Custom Livery');
    assert.ok(fs.existsSync(path.join('server/gos/catalogue', d.previewFile)));
  }
});
test('PDF matching respects category, body generation and wheelbase without substituting absent vans', () => {
  const standard = seed.liveries.find(l => l.name === 'Standard GOS').id;
  const old = {
    group: 'Ford Transit Custom 2018–2023',
    model: 'Ford Transit Custom',
    year: 2020,
    wheelbase: 'SWB',
    liveryId: standard
  };
  assert.equal(matchDesign(old, designs).pageNumber, 1);
  assert.equal(matchDesign({
    ...old,
    wheelbase: 'LWB'
  }, designs).pageNumber, 2);
  assert.equal(matchDesign({
    ...old,
    group: 'Ford Transit Connect 2024 onwards',
    model: 'Ford Transit Connect',
    year: 2025,
    wheelbase: 'LWB'
  }, designs).pageNumber, 5);
  assert.equal(matchDesign({
    ...old,
    group: 'Volkswagen Transporter T6.1 2019–2024',
    model: 'Volkswagen Transporter'
  }, designs), null);
  assert.equal(matchDesign({
    ...old,
    liveryId: seed.liveries.find(l => l.name === 'Mechanical Electrical').id,
    wheelbase: 'LWB'
  }, designs), null);
});
test('repair PDF preserves the original page dimensions and artwork and adds vector boxes', async () => {
  const original = fs.readFileSync('server/gos/catalogue/GOSVans.pdf'),
    source = await PDFDocument.load(original);
  const bytes = await pagePdf(original, 1, [{
    x: .4,
    y: .4,
    w: .2,
    h: .15
  }]);
  const repaired = await PDFDocument.load(bytes);
  assert.equal(repaired.getPageCount(), 1);
  assert.deepEqual(repaired.getPage(0).getSize(), source.getPage(0).getSize());
  assert.ok(bytes.length > 1000);
  assert.equal(fs.readFileSync('server/gos/catalogue/GOSVans.pdf').equals(original), true);
  fs.mkdirSync('outputs/gos-pdf-check', {
    recursive: true
  });
  fs.writeFileSync('outputs/gos-pdf-check/repair.pdf', bytes);
});
test('repair coordinates correctly translate top-left preview boxes for page rotation', async () => {
  const doc = await PDFDocument.create(),
    page = doc.addPage([600, 400]),
    box = {
      x: .1,
      y: .2,
      w: .3,
      h: .4
    };
  assert.deepEqual(pdfBox(page, box), {
    x: 60,
    y: 160,
    width: 180,
    height: 160
  });
  page.setRotation(degrees(90));
  const b = pdfBox(page, box);
  assert.equal(b.x, 120);
  assert.equal(b.y, 40);
  assert.equal(b.width, 240);
  assert.equal(b.height, 120);
});
