import {useEffect,useId,useRef,useState} from 'react';
import {MODULE_META,isPortalWrite} from './portal-model.mjs';

export function PortalHeading({active}) {
  const meta=MODULE_META[active];
  if(!meta) return null;
  return <div className="portal-heading"><div><p className="portal-breadcrumb">{meta[2]} <span aria-hidden="true">/</span> {meta[0]}</p><h1>{meta[0]}</h1><p>{meta[1]}</p></div><span className="portal-test-pill">Test workspace</span></div>;
}

export function PortalActionMenu({children,label='More actions'}) {
  const ref=useRef(null);
  useEffect(()=>{
    const close=e=>{if(ref.current && !ref.current.contains(e.target))ref.current.open=false;};
    const key=e=>{if(e.key==='Escape' && ref.current?.open){ref.current.open=false;ref.current.querySelector('summary').focus();}};
    document.addEventListener('pointerdown',close);document.addEventListener('keydown',key);
    return ()=>{document.removeEventListener('pointerdown',close);document.removeEventListener('keydown',key);};
  },[]);
  return <details className="portal-action-menu" ref={ref}><summary>{label}<span aria-hidden="true"> ⋯</span></summary><div className="portal-action-popover" onClick={e=>{if(e.target.closest('button')&&!e.target.closest('button').disabled)ref.current.open=false;}}>{children}</div></details>;
}

export function PortalDrawerPanel({children,className='',onClose,label='Details',...props}) {
  const ref=useRef(null),closeRef=useRef(onClose);
  closeRef.current=onClose;
  useEffect(()=>{
    const previous=document.activeElement,oldOverflow=document.body.style.overflow;
    const panel=ref.current;
    const focusables=()=>[...panel.querySelectorAll('button,input,select,textarea,a[href],summary,[tabindex="0"]')].filter(e=>!e.disabled&&e.getClientRects().length);
    document.body.style.overflow='hidden';
    (focusables()[0]||panel).focus();
    const key=e=>{
      if([...document.querySelectorAll('[data-portal-drawer]')].at(-1)!==panel)return;
      if(e.key==='Escape'){e.preventDefault();e.stopPropagation();closeRef.current?.();}
      if(e.key==='Tab'){
        const list=focusables(),first=list[0],last=list.at(-1);
        if(!first){e.preventDefault();panel.focus();return;}
        if(e.shiftKey && (document.activeElement===first || document.activeElement===panel)){e.preventDefault();last.focus();}
        else if(!e.shiftKey && document.activeElement===last){e.preventDefault();first.focus();}
      }
    };
    panel.addEventListener('keydown',key);
    return ()=>{panel.removeEventListener('keydown',key);document.body.style.overflow=oldOverflow;if(previous?.isConnected)previous.focus();};
  },[]);
  return <div {...props} ref={ref} className={'portal-drawer '+className} data-portal-drawer role="dialog" aria-modal="true" aria-label={label} tabIndex={-1} onClick={e=>e.stopPropagation()}>{children}</div>;
}

export function PortalSteps({steps,active,onChange,maxStep=steps.length-1}) {
  return <nav className="portal-steps" aria-label="Invoice workflow">{steps.map((step,i)=><button type="button" key={step} className={i===active?'active':i<active?'complete':''} aria-current={i===active?'step':undefined} disabled={i>maxStep} onClick={()=>onChange?.(i)}><span>{i+1}</span><strong>{step}</strong></button>)}</nav>;
}

export function installPortalFeedback() {
  const original=window.fetch.bind(window);
  let sequence=0;
  window.fetch=async(input,options)=>{
    if(!isPortalWrite(input,options))return original(input,options);
    const id=++sequence;
    const emit=(state,message)=>window.dispatchEvent(new CustomEvent('portal-request',{detail:{id,state,message}}));
    emit('pending','Saving changes…');
    try {
      const response=await original(input,options);
      if(response.ok)emit('success',String(options?.method||input?.method||'POST').toUpperCase()==='DELETE'?'Removed from the test workspace':'Changes saved');
      else {
        const payload=await response.clone().json().catch(()=>({}));
        emit('error',payload.error || 'Changes could not be saved. Please try again.');
      }
      return response;
    } catch(error){emit('error','Connection lost. Your changes have not been confirmed as saved.');throw error;}
  };
}

export function PortalFeedback() {
  const [state,setState]=useState(null),pending=useRef(new Set()),failure=useRef(null),timer=useRef(null);
  const id=useId();
  useEffect(()=>{
    const receive=({detail})=>{
      window.clearTimeout(timer.current);
      if(detail.state==='pending'){pending.current.add(detail.id);if(pending.current.size===1)failure.current=null;}
      else pending.current.delete(detail.id);
      if(detail.state==='error')failure.current=detail;
      setState(failure.current || (pending.current.size?{state:'pending',message:'Saving changes…'}:detail));
      if(!pending.current.size && !failure.current)timer.current=window.setTimeout(()=>setState(null),4500);
    };
    window.addEventListener('portal-request',receive);
    return ()=>{window.removeEventListener('portal-request',receive);window.clearTimeout(timer.current);};
  },[]);
  return state?<div className={'portal-feedback is-'+state.state} role={state.state==='error'?'alert':'status'} aria-live={state.state==='error'?'assertive':'polite'} id={id}><span className="portal-feedback-symbol" aria-hidden="true">{state.state==='pending'?'◌':state.state==='error'?'!':'✓'}</span><span>{state.message}</span>{state.state!=='pending'?<button type="button" aria-label="Dismiss status" onClick={()=>{setState(null);failure.current=null;}}>×</button>:null}</div>:null;
}
