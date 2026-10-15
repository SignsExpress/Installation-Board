import React, { useState } from 'react';
export default function PdfUpload({
  liveries,
  documents,
  designs,
  onUpload,
  vehicle
}) {
  const [replaceId, setReplaceId] = useState('');
  const target = documents.find(d => d.id === replaceId);
  const [file, setFile] = useState(null),
    [pages, setPages] = useState([]),
    [busy, setBusy] = useState(false),
    [message, setMessage] = useState(''),
    [error, setError] = useState('');
  const labelKey = title => String(title || '').normalize('NFKC').toLowerCase().replace(/[^a-z0-9]/g, '');
  async function prepare(input) {
    if (!input) return;
    setError('');
    setPages([]);
    setFile(null);
    if (input.size > 22 * 1024 * 1024) {
      setError('Choose a PDF under 22MB.');
      return;
    }
    setBusy(true);
    let pdf;
    try {
      const bytes = await input.arrayBuffer(),
        lib = await import('pdfjs-dist/build/pdf.mjs');
      lib.GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/build/pdf.worker.min.mjs', import.meta.url).href;
      pdf = await lib.getDocument({
        data: new Uint8Array(bytes.slice(0)),
        isEvalSupported: false
      }).promise;
      if (vehicle && pdf.numPages !== 1) throw Error('Choose a one-page PDF for this van.');
      if (pdf.numPages > 100) throw Error('Choose a PDF with at most 100 pages.');
      const result = [];
      for (let i = 1; i <= pdf.numPages; i++) {
        setMessage('Reading page ' + i + ' of ' + pdf.numPages + '…');
        const page = await pdf.getPage(i),
          viewport = page.getViewport({
            scale: 1
          }),
          canvas = document.createElement('canvas'),
          renderView = page.getViewport({
            scale: Math.min(1, 1400 / Math.max(viewport.width, viewport.height))
          });
        canvas.width = Math.ceil(renderView.width);
        canvas.height = Math.ceil(renderView.height);
        await page.render({
          canvasContext: canvas.getContext('2d'),
          viewport: renderView
        }).promise;
        const text = await page.getTextContent(),
          title = text.items.map(t => t.str).join(' ').trim();
        const group = pdf.numPages === 19 ? i <= 9 ? 'Standard GOS' : i <= 14 ? 'Mechanical Electrical' : i <= 16 ? 'Electrical' : 'Custom Livery' : '';
        result.push({
          pageNumber: i,
          title: vehicle ? vehicle.reg + ' - ' + vehicle.model : title.slice(0, 500) || 'Page ' + i,
          liveryId: vehicle?.liveryId || designs.find(d => d.documentId === replaceId && labelKey(d.title) === labelKey(title.slice(0, 500)))?.liveryId || liveries.find(l => l.name === group)?.id || '',
          previewDataUrl: canvas.toDataURL('image/jpeg', .86)
        });
        page.cleanup();
      }
      let binary = '';
      for (let i = 0; i < bytes.byteLength; i += 32768) binary += String.fromCharCode(...new Uint8Array(bytes.slice(i, i + 32768)));
      setFile({
        name: input.name,
        dataUrl: 'data:application/pdf;base64,' + btoa(binary)
      });
      setPages(result);
      setMessage(vehicle ? 'Check the drawing, then save.' : 'Check each van type and category, then save.');
    } catch (e) {
      setError(e.message);
    } finally {
      await pdf?.destroy();
      setBusy(false);
    }
  }
  async function save() {
    setBusy(true);
    setError('');
    try {
      if (pages.some(p => !p.liveryId)) throw Error('Choose a livery category for each page.');
      await onUpload({
        ...file,
        ...(vehicle ? {vehicleId: vehicle.id, vehicleVersion: vehicle.version, vehicle} : {}),
        replaceId: target?.id,
        revision: target?.revision,
        pages
      });
      setFile(null);
      setPages([]);
      setMessage('PDF saved.');
    } catch (e) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  }
  return <section className="gos-pdf-import"><h3>{vehicle ? "Replace this van’s PDF" : "Add or replace a PDF"}</h3>{!vehicle && <label>What do you want to do?<select value={replaceId} disabled={busy || !!file} onChange={e => { setReplaceId(e.target.value); setMessage(''); }}><option value="">Upload a new PDF</option>{documents.map(d => <option key={d.id} value={d.id}>Replace {d.name}</option>)}</select></label>}{vehicle && <p>Choose a one-page PDF. Only this van will change.</p>}{target && <p>Replacing {target.name}. Check the van labels match. Old orders keep their original PDF.</p>}<label className="gos-upload-button">Choose PDF<input type="file" accept="application/pdf,.pdf" disabled={busy} onChange={e => prepare(e.target.files[0])} /></label>{message && <p role="status">{message}</p>}{error && <p className="gos-error" role="alert">{error}</p>}{pages.length > 0 && <><div className="gos-pdf-import-pages">{pages.map((p, i) => <article key={p.pageNumber}><img src={p.previewDataUrl} alt={'PDF page ' + p.pageNumber} /><strong>{vehicle ? vehicle.reg : "Page " + p.pageNumber}</strong>{!vehicle && <><label>Van type<input value={p.title} onChange={e => setPages(rows => rows.map((r, j) => j === i ? {
              ...r,
              title: e.target.value
            } : r))} /></label><label>Livery category<select value={p.liveryId} onChange={e => setPages(rows => rows.map((r, j) => j === i ? {
              ...r,
              liveryId: e.target.value
            } : r))}><option value="">Choose category</option>{liveries.map(l => <option key={l.id} value={l.id}>{l.name}</option>)}</select></label></>}</article>)}</div><button type="button" className="gos-primary" disabled={busy} onClick={save}>{busy ? 'Saving PDF…' : vehicle ? 'Save PDF and van' : target ? 'Replace PDF' : 'Save PDF'}</button><button type="button" disabled={busy} onClick={() => {
        setPages([]);
        setFile(null);
        setMessage('');
      }}>Cancel upload</button></>}</section>;
}
