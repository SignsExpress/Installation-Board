const {
  PDFDocument,
  rgb,
  StandardFonts,
  degrees
} = require('pdf-lib');
function family(text) {
  const t = text.toLowerCase();
  if (t.includes('dropside') || t.includes('pickup') || t.includes('pick up')) return 'dropside';
  if (t.includes('courier')) return 'courier';
  if (t.includes('connect')) return 'connect';
  if (t.includes('custom')) return 'custom';
  if (t.includes('transit') && (t.includes('panel') || t.includes('full size'))) return 'panel';
  return '';
}
function wb(text) {
  const t = text.toLowerCase();
  if (/double\s*cab/.test(t)) return 'double';
  if (/single\s*cab/.test(t)) return 'single';
  if (/\bl1\b|\bswb\b/.test(t)) return 'short';
  if (/\bl3\b/.test(t)) return 'long';
  if (/\blwb\b/.test(t)) return 'long';
  if (/\bmwb\b/.test(t)) return 'medium';
  return '';
}
function matchDesign(vehicle, designs) {
  const candidates = designs.filter(d => d.liveryId === vehicle.liveryId);
  if (vehicle.designId) return candidates.find(d => d.id === vehicle.designId) || null;
  const text = [vehicle.group, vehicle.model, vehicle.wheelbase].join(' '),
    f = family(text);
  if (!f) return null;
  let found = candidates.filter(d => family(d.title) === f);
  const yearMatch = String(vehicle.group || '').match(/(20\d{2})/),
    year = yearMatch ? Number(yearMatch[1]) : Number(vehicle.year);
  if (yearMatch) {
    const exactGeneration = found.filter(d => Number(d.title.match(/20\d{2}/)?.[0]) === Number(yearMatch[1]));
    if (exactGeneration.length) found = exactGeneration;
  }
  if (year) found = found.filter(d => {
    const m = d.title.match(/(20\d{2})(?:\s*[-–]\s*(\d{2,4}))?/);
    return !m || Number(m[1]) <= year && (!m[2] || year <= Number(m[2].length === 2 ? '20' + m[2] : m[2]));
  });
  const wheel = wb(vehicle.wheelbase) || wb(vehicle.model + ' ' + vehicle.group);
  if (wheel) found = found.filter(d => !wb(d.title) || wb(d.title) === wheel);
  return found.length === 1 ? found[0] : null;
}
function pdfBox(page, b) {
  const c = page.getCropBox(),
    r = (page.getRotation().angle % 360 + 360) % 360;
  switch (r) {
    case 90:
      return {
        x: c.x + b.y * c.width,
        y: c.y + b.x * c.height,
        width: b.h * c.width,
        height: b.w * c.height
      };
    case 180:
      return {
        x: c.x + (1 - b.x - b.w) * c.width,
        y: c.y + b.y * c.height,
        width: b.w * c.width,
        height: b.h * c.height
      };
    case 270:
      return {
        x: c.x + (1 - b.y - b.h) * c.width,
        y: c.y + (1 - b.x - b.w) * c.height,
        width: b.h * c.width,
        height: b.w * c.height
      };
    default:
      return {
        x: c.x + b.x * c.width,
        y: c.y + (1 - b.y - b.h) * c.height,
        width: b.w * c.width,
        height: b.h * c.height
      };
  }
}
async function pagePdf(bytes, pageNumber, boxes = []) {
  const original = await PDFDocument.load(bytes);
  if (pageNumber < 1 || pageNumber > original.getPageCount()) throw Error('PDF page not found.');
  const out = await PDFDocument.create(),
    [page] = await out.copyPages(original, [pageNumber - 1]);
  out.addPage(page);
  const font = await out.embedFont(StandardFonts.HelveticaBold);
  boxes.forEach((b, i) => {
    const p = pdfBox(page, b),
      size = Math.max(12, page.getWidth() / 100);
    page.drawRectangle({
      ...p,
      borderColor: rgb(.92, .06, .1),
      borderWidth: Math.max(2, page.getWidth() / 400)
    });
    page.drawText(String(i + 1), {
      x: p.x + size * .3,
      y: p.y + p.height - size * 1.2,
      size,
      font,
      color: rgb(.92, .06, .1),
      rotate: degrees(0)
    });
  });
  out.setTitle(boxes.length ? 'GOS vehicle repair areas' : 'GOS vehicle livery');
  return Buffer.from(await out.save());
}
module.exports = {
  matchDesign,
  pagePdf,
  pdfBox
};
