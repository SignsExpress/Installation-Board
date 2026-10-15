import React, { useState } from 'react';
export default function PdfUpload({
  liveries,
  onUpload
}) {
  const [file, setFile] = useState(null),
    [pages, setPages] = useState([]),
    [busy, setBusy] = useState(false),
    [message, setMessage] = useState(''),
    [error, setError] = useState('');
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
          title: title.slice(0, 500) || 'Page ' + i,
          liveryId: liveries.find(l => l.name === group)?.id || '',
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
      setMessage('Check the van label and livery category for every page, then save the PDF.');
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
        pages
      });
      setFile(null);
      setPages([]);
      setMessage('PDF uploaded. Its original pages are available for orders and repairs.');
    } catch (e) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  }
  return <section className="gos-pdf-import"><h3>Upload a livery PDF</h3><p>The original PDF is retained. Page snapshots are generated only for previews and repair-box positioning.</p><label className="gos-upload-button">Choose PDF<input type="file" accept="application/pdf,.pdf" disabled={busy} onChange={e => prepare(e.target.files[0])} /></label>{message && <p role="status">{message}</p>}{error && <p className="gos-error" role="alert">{error}</p>}{pages.length > 0 && <><div className="gos-pdf-import-pages">{pages.map((p, i) => <article key={p.pageNumber}><img src={p.previewDataUrl} alt={'PDF page ' + p.pageNumber} /><strong>Page {p.pageNumber}</strong><label>Van type<input value={p.title} onChange={e => setPages(rows => rows.map((r, j) => j === i ? {
              ...r,
              title: e.target.value
            } : r))} /></label><label>Livery category<select value={p.liveryId} onChange={e => setPages(rows => rows.map((r, j) => j === i ? {
              ...r,
              liveryId: e.target.value
            } : r))}><option value="">Choose category</option>{liveries.map(l => <option key={l.id} value={l.id}>{l.name}</option>)}</select></label></article>)}</div><button className="gos-primary" disabled={busy} onClick={save}>{busy ? 'Saving PDF…' : 'Save PDF and page categories'}</button><button disabled={busy} onClick={() => {
        setPages([]);
        setFile(null);
        setMessage('');
      }}>Cancel upload</button></>}</section>;
}
