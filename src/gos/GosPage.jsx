import React, { useEffect, useRef, useState, lazy, Suspense } from 'react';
const PdfUpload = lazy(() => import('./PdfUpload'));
import './gos.css';
const blank = {
  reg: '',
  name: '',
  division: '',
  year: '',
  make: '',
  model: '',
  group: '',
  wheelbase: '',
  colour: '',
  liveryId: '',
  designId: '',
  notes: ''
};
async function api(route, body, method = body ? 'POST' : 'GET') {
  const r = await fetch('/api/gos' + route, {
    method,
    headers: body ? {
      'Content-Type': 'application/json'
    } : undefined,
    body: body ? JSON.stringify(body) : undefined
  });
  const data = await r.json();
  if (!r.ok) {
    if (r.status === 401) window.dispatchEvent(new Event('gos-session-expired'));
    const e = Error(data.error || 'Request failed.');
    e.status = r.status;
    throw e;
  }
  return data;
}
function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(Error('Could not load an image.'));
    img.src = src;
  });
}
async function imageData(file) {
  if (!['image/jpeg', 'image/png', 'image/webp'].includes(file.type)) throw Error('Choose a PNG, JPEG or WebP image.');
  if (file.size > 15 * 1024 * 1024) throw Error('Choose an image under 15MB.');
  const url = URL.createObjectURL(file);
  try {
    const img = await loadImage(url),
      scale = Math.min(1, 1800 / img.width, 1800 / img.height),
      canvas = document.createElement('canvas');
    canvas.width = Math.round(img.width * scale);
    canvas.height = Math.round(img.height * scale);
    canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
    return canvas.toDataURL('image/jpeg', .9);
  } finally {
    URL.revokeObjectURL(url);
  }
}
async function markedPhoto(photo, boxes) {
  const img = await loadImage(photo.url),
    canvas = document.createElement('canvas'),
    scale = Math.min(1, 1800 / img.width);
  canvas.width = Math.round(img.width * scale);
  canvas.height = Math.round(img.height * scale);
  const c = canvas.getContext('2d');
  c.drawImage(img, 0, 0, canvas.width, canvas.height);
  c.lineWidth = Math.max(3, canvas.width / 250);
  c.strokeStyle = '#e12433';
  c.font = 'bold ' + Math.max(18, canvas.width / 55) + 'px Arial';
  boxes.forEach((b, i) => {
    const x = b.x * canvas.width,
      y = b.y * canvas.height;
    c.strokeRect(x, y, b.w * canvas.width, b.h * canvas.height);
    c.fillStyle = '#e12433';
    c.fillRect(x, y, 32, 30);
    c.fillStyle = 'white';
    c.fillText(String(i + 1), x + 8, y + 23);
  });
  return {
    photoId: photo.id,
    boxes,
    dataUrl: canvas.toDataURL('image/jpeg', .9)
  };
}
function Modal({
  title,
  onClose,
  children
}) {
  const ref = useRef(null);
  useEffect(() => {
    const old = document.activeElement,
      overflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    ref.current?.focus();
    return () => {
      document.body.style.overflow = overflow;
      old?.focus?.();
    };
  }, []);
  return <div className="gos-backdrop"><section className="gos-modal" role="dialog" aria-modal="true" aria-label={title} tabIndex={-1} ref={ref} onKeyDown={e => {
      if (e.key === 'Escape') onClose();
      if (e.key === 'Tab') {
        const list = Array.from(ref.current.querySelectorAll('button,input,select,textarea,a[href]')).filter(el => !el.disabled && el.getClientRects().length);
        if (e.shiftKey && (document.activeElement === list[0] || document.activeElement === ref.current)) {
          e.preventDefault();
          list.at(-1)?.focus();
        } else if (!e.shiftKey && document.activeElement === list.at(-1)) {
          e.preventDefault();
          list[0]?.focus();
        }
      }
    }}><header><h2>{title}</h2><button type="button" onClick={onClose} aria-label="Close dialog">×</button></header>{children}</section></div>;
}
function Fields({
  value,
  onChange,
  liveries,
  divisions,
  full = true,
  designs = []
}) {
  const update = (key, val) => onChange({
    ...value,
    [key]: val
  });
  return <div className="gos-fields">{[['reg', 'Registration'], ['make', 'Make'], ['model', 'Model'], ['wheelbase', 'Wheelbase / roof'], ['colour', 'Colour'], ...(full ? [['name', 'Driver / assigned to'], ['division', 'Division'], ['year', 'Year'], ['group', 'Vehicle group']] : [])].map(([key, label]) => <label key={key}>{label}<input list={key === 'division' ? 'gos-divisions' : undefined} required={['reg', 'model'].includes(key)} value={value[key] || ''} onChange={e => update(key, e.target.value)} /></label>)}<datalist id="gos-divisions">{divisions.map(d => <option key={d}>{d}</option>)}</datalist><label>Livery<select required value={value.liveryId} onChange={e => onChange({
        ...value,
        liveryId: e.target.value,
        designId: ''
      })}><option value="">Choose a livery</option>{liveries.map(l => <option key={l.id} value={l.id}>{l.name}</option>)}</select></label>{designs.length > 0 && <label className="gos-wide">PDF van design<select value={value.designId || ''} onChange={e => update('designId', e.target.value)}><option value="">Select matching van / wheelbase</option>{designs.filter(d => d.liveryId === value.liveryId).map(d => <option key={d.id} value={d.id}>Page {d.pageNumber} — {d.title}</option>)}</select></label>}{full && <label className="gos-wide">Vehicle notes<textarea rows={3} value={value.notes} onChange={e => update('notes', e.target.value)} /></label>}</div>;
}
function RepairImage({
  photo,
  boxes,
  onChange
}) {
  const ref = useRef(null),
    [start, setStart] = useState(null),
    [draft, setDraft] = useState(null);
  function point(e) {
    const r = ref.current.getBoundingClientRect();
    return {
      x: Math.max(0, Math.min(1, (e.clientX - r.left) / r.width)),
      y: Math.max(0, Math.min(1, (e.clientY - r.top) / r.height))
    };
  }
  function rect(a, b) {
    return {
      x: Math.min(a.x, b.x),
      y: Math.min(a.y, b.y),
      w: Math.abs(a.x - b.x),
      h: Math.abs(a.y - b.y)
    };
  }
  return <article className="gos-repair-photo"><div className="gos-repair-surface" ref={ref} onPointerDown={e => {
      if (e.button !== 0) return;
      e.preventDefault();
      e.currentTarget.setPointerCapture(e.pointerId);
      setStart(point(e));
    }} onPointerMove={e => {
      if (start) setDraft(rect(start, point(e)));
    }} onPointerUp={e => {
      if (start) {
        const r = rect(start, point(e));
        if (r.w > .005 && r.h > .005) onChange([...boxes, r]);
      }
      setStart(null);
      setDraft(null);
    }} onPointerCancel={() => {
      setStart(null);
      setDraft(null);
    }}><img src={photo.url} alt={photo.name || 'Vehicle repair reference'} draggable={false} /><svg viewBox="0 0 1000 1000" preserveAspectRatio="none" aria-label="Marked repair areas">{[...boxes, ...(draft ? [draft] : [])].map((b, i) => <g key={i}><rect x={b.x * 1000} y={b.y * 1000} width={b.w * 1000} height={b.h * 1000} fill="#ed243322" stroke="#ed2433" strokeWidth="4" vectorEffect="non-scaling-stroke" /><text x={b.x * 1000 + 8} y={b.y * 1000 + 30} fill="#cf1122" fontSize="30" fontWeight="bold">{i + 1}</text></g>)}</svg></div><div className="gos-photo-tools"><span>{boxes.length} repair {boxes.length === 1 ? 'area' : 'areas'}</span><button type="button" disabled={!boxes.length} onClick={() => onChange(boxes.slice(0, -1))}>Undo last box</button><button type="button" disabled={!boxes.length} onClick={() => onChange([])}>Clear boxes</button></div></article>;
}
function Order({
  kind,
  initialId,
  register,
  onClose,
  onSent
}) {
  const [mode, setMode] = useState(initialId ? 'existing' : 'existing'),
    [existingId, setExistingId] = useState(initialId || ''),
    [vehicle, setVehicle] = useState({
      ...blank
    }),
    [liveryId, setLiveryId] = useState(register.vehicles.find(v => v.id === initialId)?.liveryId || ''),
    [designId, setDesignId] = useState(register.vehicles.find(v => v.id === initialId)?.matchedDesign?.id || ''),
    [contact, setContact] = useState(''),
    [email, setEmail] = useState(''),
    [notes, setNotes] = useState(''),
    [boxes, setBoxes] = useState({}),
    [error, setError] = useState(''),
    [sending, setSending] = useState(false);
  const requestId = useRef(crypto.randomUUID()),
    existing = register.vehicles.find(v => v.id === existingId),
    livery = register.liveries.find(l => l.id === liveryId);
  const design = (register.designs || []).find(d => d.id === designId && d.liveryId === liveryId);
  const photos = design ? [{
    id: design.id,
    url: design.previewUrl,
    name: design.title
  }] : mode === 'existing' && existing?.liveryId === liveryId && existing?.photos?.length ? existing.photos : livery?.photos || [];
  const divisions = [...new Set(register.vehicles.map(v => v.division).filter(Boolean))];
  async function submit(e) {
    e.preventDefault();
    setError('');
    if (!photos.length) {
      setError('Select the matching PDF design, or upload a reference document first.');
      return;
    }
    if (kind === 'repair' && !photos.some(p => boxes[p.id]?.length)) {
      setError('Draw at least one box on a photo to show the repair area.');
      return;
    }
    setSending(true);
    try {
      const annotations = kind === 'repair' ? design ? [{
        designId: design.id,
        boxes: boxes[design.id] || []
      }] : await Promise.all(photos.filter(p => boxes[p.id]?.length).map(p => markedPhoto(p, boxes[p.id]))) : [];
      const result = await api('/orders', {
        requestId: requestId.current,
        kind,
        existingId: mode === 'existing' ? existingId : null,
        vehicle: {
          ...vehicle,
          liveryId
        },
        liveryId,
        designId: design?.id || '',
        contact,
        email,
        notes,
        annotations
      });
      onSent(result.id);
    } catch (e) {
      setError(e.message);
    } finally {
      setSending(false);
    }
  }
  return <Modal title={kind === 'repair' ? 'Order a repair' : 'Order vehicle livery'} onClose={() => {
    if (!sending) onClose();
  }}><form onSubmit={submit}><p className="gos-muted">Choose the vehicle and livery, then review the details before sending the order to Signs Express.</p><div className="gos-segment"><button type="button" aria-pressed={mode === 'existing'} onClick={() => {
          setMode('existing');
          setBoxes({});
        }}>Existing registration</button><button type="button" aria-pressed={mode === 'new'} onClick={() => {
          setMode('new');
          setExistingId('');
          setBoxes({});
        }}>New registration</button></div>{mode === 'existing' ? <><label>Vehicle<select required value={existingId} onChange={e => {
            const v = register.vehicles.find(v => v.id === e.target.value);
            setExistingId(e.target.value);
            setLiveryId(v?.liveryId || '');
            setDesignId(v?.matchedDesign?.id || '');
            setBoxes({});
          }}><option value="">Choose a registration</option>{register.vehicles.map(v => <option key={v.id} value={v.id}>{v.reg} — {v.model}</option>)}</select></label>{existing && <p className="gos-summary">{existing.make} · {existing.model} · {existing.wheelbase || 'Wheelbase not recorded'} · {existing.colour || 'Colour not recorded'}<br />{existing.division} · {existing.name}</p>}<label>Livery<select required value={liveryId} onChange={e => {
            setLiveryId(e.target.value);
            setDesignId('');
            setBoxes({});
          }}><option value="">Choose a livery</option>{register.liveries.map(l => <option key={l.id} value={l.id}>{l.name}</option>)}</select></label></> : <Fields full={false} value={{
        ...vehicle,
        liveryId
      }} onChange={v => {
        setVehicle(v);
        if (v.liveryId !== liveryId) {
          setLiveryId(v.liveryId);
          setBoxes({});
        }
      }} liveries={register.liveries} divisions={divisions} />}<label className="gos-design-choice">PDF van design<select value={design?.id || ''} onChange={e => {
          setDesignId(e.target.value);
          setBoxes({});
        }}><option value="">Choose the matching van / wheelbase</option>{(register.designs || []).filter(d => d.liveryId === liveryId).map(d => <option key={d.id} value={d.id}>Page {d.pageNumber} — {d.title}</option>)}</select></label>{design && <p><a href={design.pdfUrl} target="_blank" rel="noreferrer">Open original PDF page</a> · <a href={design.documentUrl} target="_blank" rel="noreferrer">Open complete source PDF</a></p>}<section className="gos-order-images"><h3>{kind === 'repair' ? 'Mark the areas that need repairing' : 'Livery / vehicle reference'}</h3>{kind === 'repair' && <p>Drag boxes over the damaged graphics. You can mark several areas on each image; the original PDF page with your repair boxes will be attached to the order email.</p>}{!photos.length ? <p className="gos-empty">Choose a matching PDF design above. If the exact van is not listed, upload the correct PDF rather than selecting a different vehicle.</p> : photos.map(p => kind === 'repair' ? <RepairImage key={p.id} photo={p} boxes={boxes[p.id] || []} onChange={b => setBoxes({
          ...boxes,
          [p.id]: b
        })} /> : <img className="gos-order-image" key={p.id} src={p.url} alt={p.name || livery?.name} />)}</section><div className="gos-fields"><label>Your name<input required value={contact} onChange={e => setContact(e.target.value)} /></label><label>Contact email<input required type="email" value={email} onChange={e => setEmail(e.target.value)} /></label><label className="gos-wide">Order notes / reference<textarea rows={4} value={notes} onChange={e => setNotes(e.target.value)} placeholder="Purchase order, required date, delivery or repair instructions" /></label></div>{error && <p role="alert" className="gos-error">{error}</p>}{!register.orderingReady && <p className="gos-error">The order inbox has not been configured yet.</p>}<footer><button type="button" disabled={sending} onClick={onClose}>Cancel</button><button className="gos-primary" disabled={sending || !register.orderingReady}>{sending ? 'Sending order…' : 'Send ' + (kind === 'repair' ? 'repair' : 'livery') + ' order'}</button></footer></form></Modal>;
}
export default function GosPage() {
  const [register, setRegister] = useState(null),
    [loading, setLoading] = useState(true),
    [password, setPassword] = useState(''),
    [error, setError] = useState(''),
    [busy, setBusy] = useState(false),
    [search, setSearch] = useState(''),
    [division, setDivision] = useState(''),
    [liveryFilter, setLiveryFilter] = useState(''),
    [tab, setTab] = useState('vehicles'),
    [edit, setEdit] = useState(null),
    [order, setOrder] = useState(null),
    [notice, setNotice] = useState(''),
    [newLivery, setNewLivery] = useState('');
  async function refresh() {
    const r = await api('/register');
    setRegister(r);
  }
  useEffect(() => {
    const expired = () => {
      setRegister(null);
      setEdit(null);
      setOrder(null);
    };
    window.addEventListener('gos-session-expired', expired);
    refresh().catch(e => {
      if (e.status !== 401) setError(e.message);
    }).finally(() => setLoading(false));
    return () => window.removeEventListener('gos-session-expired', expired);
  }, []);
  async function login(e) {
    e.preventDefault();
    setBusy(true);
    setError('');
    try {
      await api('/login', {
        password
      });
      setPassword('');
      await refresh();
    } catch (e) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  }
  async function save(e) {
    e.preventDefault();
    setBusy(true);
    setError('');
    try {
      await api('/vehicles' + (edit.id ? '/' + edit.id : ''), edit, edit.id ? 'PATCH' : 'POST');
      setEdit(null);
      await refresh();
      setNotice('Vehicle saved.');
    } catch (e) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  }
  async function upload(kind, id, file) {
    if (!file) return;
    setBusy(true);
    setError('');
    try {
      const dataUrl = await imageData(file);
      await api('/' + kind + '/' + id + '/photos', {
        dataUrl,
        name: file.name
      });
      await refresh();
      setNotice('Photo uploaded.');
    } catch (e) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  }
  if (loading) return <div className="gos-app"><p className="gos-loading">Loading GOS Vehicle Liveries…</p></div>;
  if (!register) return <div className="gos-app gos-login"><form onSubmit={login}><div className="gos-wordmark">GOS</div><p className="gos-eyebrow">VEHICLE LIVERIES</p><h1>Your fleet, every livery.</h1><p>View your vehicle register, request new graphics and mark repairs.</p><label>Portal password<input type="password" autoComplete="current-password" required value={password} onChange={e => setPassword(e.target.value)} /></label>{error && <p role="alert" className="gos-error">{error}</p>}<button className="gos-primary" disabled={busy}>{busy ? 'Signing in…' : 'Open vehicle portal'}</button><small>Managed by Signs Express Central Lancashire & Southport</small></form></div>;
  const divisions = [...new Set(register.vehicles.map(v => v.division).filter(Boolean))].sort();
  const vehicles = register.vehicles.filter(v => (!division || v.division === division) && (!liveryFilter || v.liveryId === liveryFilter) && (!search || [v.reg, v.name, v.division, v.model, v.group, register.liveries.find(l => l.id === v.liveryId)?.name].join(' ').toLowerCase().includes(search.toLowerCase()) || v.reg.replace(/\s/g, '').toLowerCase().includes(search.replace(/\s/g, '').toLowerCase())));
  return <div className="gos-app"><header className="gos-top"><div><span className="gos-wordmark">GOS</span><span>Vehicle Liveries</span></div><nav><button onClick={() => setOrder({
          kind: 'livery'
        })} className="gos-primary">Order livery</button><button onClick={() => setOrder({
          kind: 'repair'
        })}>Order repair</button><button onClick={async () => {
          await api('/logout', {});
          setRegister(null);
        }}>Sign out</button></nav></header><main>{register.previewMode && <div className="gos-info">Test preview — orders are captured locally. No emails are sent.</div>}<div className="gos-heading"><div><p className="gos-eyebrow">GOS FLEET PORTAL</p><h1>GOS Vehicle Liveries</h1><p>Find your vehicles, check their liveries and keep the register up to date.</p></div><span className="gos-total">{register.vehicles.length} vehicles</span></div>{notice && <div className="gos-notice" role="status">{notice}<button onClick={() => setNotice('')} aria-label="Dismiss notification">×</button></div>}{error && !edit && <p className="gos-error" role="alert">{error}</p>}<div className="gos-tabs">{[['vehicles', 'Vehicle register'], ['liveries', 'Livery designs'], ['orders', 'Order history']].map(([id, label]) => <button key={id} aria-pressed={tab === id} onClick={() => setTab(id)}>{label}</button>)}</div>{tab === 'vehicles' ? <><div className="gos-filters"><label>Search<input type="search" placeholder="Registration, driver, model…" value={search} onChange={e => setSearch(e.target.value)} /></label><label>Division<select value={division} onChange={e => setDivision(e.target.value)}><option value="">All divisions</option>{divisions.map(d => <option key={d}>{d}</option>)}</select></label><label>Livery<select value={liveryFilter} onChange={e => setLiveryFilter(e.target.value)}><option value="">All liveries</option>{register.liveries.map(l => <option value={l.id} key={l.id}>{l.name}</option>)}</select></label><button onClick={() => {
            setEdit({
              ...blank,
              liveryId: register.liveries[0]?.id
            });
            setError('');
          }}>Add vehicle</button></div><p className="gos-result-count">Showing {vehicles.length} of {register.vehicles.length} vehicles</p><div className="gos-vehicle-grid">{vehicles.map(v => {
            const l = register.liveries.find(l => l.id === v.liveryId),
              p = v.matchedDesign ? {
                url: v.matchedDesign.previewUrl
              } : v.photos?.[0];
            return <article key={v.id} className="gos-vehicle-card"><div className="gos-card-photo">{p ? <img src={p.url} alt={v.reg + ' vehicle photo'} loading="lazy" /> : <div><strong>{v.reg}</strong><span>No exact PDF match yet</span></div>}{v.matchedDesign ? <a className="gos-upload" href={v.matchedDesign.pdfUrl} target="_blank" rel="noreferrer">Open PDF</a> : <button className="gos-upload" onClick={() => {
                  setEdit({
                    ...v,
                    designId: ''
                  });
                  setError('');
                }}>Choose PDF design</button>}</div><div className="gos-card-body"><div className="gos-reg-row"><h2>{v.reg}</h2><button onClick={() => {
                    setEdit({
                      ...v,
                      designId: v.designId || v.matchedDesign?.id || ''
                    });
                    setError('');
                  }} aria-label={'Edit ' + v.reg}>Edit</button></div><p>{v.model}</p>{v.matchedDesign && <p className="gos-pdf-caption">PDF {v.matchedDesign.pageNumber}: {v.matchedDesign.title}</p>}<p className="gos-muted">{[v.wheelbase, v.colour].filter(Boolean).join(' · ') || 'Vehicle details can be added'}</p><dl><div><dt>Division</dt><dd>{v.division || 'Unallocated'}</dd></div><div><dt>Assigned to</dt><dd>{v.name || 'Unassigned'}</dd></div></dl><span className="gos-tag">{l?.name || 'No livery selected'}</span>{v.notes && <p>{v.notes}</p>}<div className="gos-card-actions"><button onClick={() => setOrder({
                    kind: 'livery',
                    initialId: v.id
                  })}>Order livery</button><button onClick={() => setOrder({
                    kind: 'repair',
                    initialId: v.id
                  })}>Order repair</button></div></div></article>;
          })}</div>{!vehicles.length && <p className="gos-empty">No vehicles match these filters.</p>}</> : tab === 'liveries' ? <><div className="gos-info">Choose a PDF page by livery category, van type and wheelbase. The text from the original page identifies the vehicle; snapshots are previews only. Orders and repairs use the original PDF page.</div><div className="gos-documents">{(register.documents || []).map(d => <a key={d.id} href={d.url} target="_blank" rel="noreferrer">Open {d.name} ({d.pageCount} pages)</a>)}</div><Suspense fallback={<p>Loading PDF importer…</p>}><PdfUpload liveries={register.liveries} onUpload={async payload => {
            await api('/documents', payload);
            await refresh();
          }} /></Suspense><form className="gos-add-livery" onSubmit={async e => {
          e.preventDefault();
          setError('');
          try {
            await api('/liveries', {
              name: newLivery
            });
            setNewLivery('');
            await refresh();
          } catch (e) {
            setError(e.message);
          }
        }}><input required aria-label="New livery name" placeholder="New livery category" value={newLivery} onChange={e => setNewLivery(e.target.value)} /><button>Add category</button></form>{register.liveries.slice().sort((a,b)=>{const names=['Standard GOS','Mechanical Electrical','Electrical','Custom Livery','No Livery'];return (names.indexOf(a.name)<0?99:names.indexOf(a.name))-(names.indexOf(b.name)<0?99:names.indexOf(b.name));}).map(l => <section className="gos-category" key={l.id}><div className="gos-category-head"><h2>{l.name}</h2><button onClick={() => {
              setTab('vehicles');
              setLiveryFilter(l.id);
              setDivision('');
              setSearch('');
            }}>View vehicles</button></div><div className="gos-pdf-grid">{(register.designs || []).filter(d => d.liveryId === l.id).map(d => <article className="gos-pdf-card" key={d.id}><a href={d.pdfUrl} target="_blank" rel="noreferrer"><img src={d.previewUrl} alt={d.title} loading="lazy" /></a><div><span className="gos-tag">Page {d.pageNumber}</span><h3>{d.title}</h3><a href={d.pdfUrl} target="_blank" rel="noreferrer">Open PDF page</a></div></article>)}</div>{!(register.designs || []).some(d => d.liveryId === l.id) && <p className="gos-empty">No PDF pages assigned to this category.</p>}</section>)} </> : <div className="gos-table-wrap"><table><thead><tr><th>Reference</th><th>Type</th><th>Registration</th><th>Livery</th><th>Date</th><th>Email status</th></tr></thead><tbody>{register.orders.slice().reverse().map(o => <tr key={o.id}><td>{o.id.slice(0, 8)}</td><td>{o.kind}</td><td>{o.reg}</td><td>{o.livery}</td><td>{new Date(o.createdAt).toLocaleString('en-GB')}</td><td>{o.status}</td></tr>)}</tbody></table>{!register.orders.length && <p className="gos-empty">No orders submitted yet.</p>}</div>}</main>{edit && <Modal title={edit.id ? 'Edit ' + edit.reg : 'Add vehicle'} onClose={() => {
      if (!busy) setEdit(null);
    }}><form onSubmit={save}><Fields value={edit} onChange={setEdit} liveries={register.liveries} divisions={divisions} designs={register.designs || []} />{error && <p role="alert" className="gos-error">{error}</p>}<footer><button type="button" onClick={() => setEdit(null)}>Cancel</button><button className="gos-primary" disabled={busy}>{busy ? 'Saving…' : 'Save vehicle'}</button></footer></form></Modal>}{order && <Order key={order.kind + (order.initialId || 'new')} {...order} register={register} onClose={() => setOrder(null)} onSent={async id => {
      setOrder(null);
      setNotice('Order ' + id.slice(0, 8) + ' sent to Signs Express.');
      await refresh();
    }} />}<footer className="gos-site-footer">GOS Vehicle Liveries · Managed by Signs Express Central Lancashire & Southport</footer></div>;
}
