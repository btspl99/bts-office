/* BTS PDF — small shared UI helpers for long-running tools */
import { esc } from './state.js';

/** A modal with a progress bar and a Cancel button. Returns { set(frac, text), close(), cancelled, onCancel }. */
export function progress(title, text = '') {
  const back = document.createElement('div'); back.className = 'modal-back';
  back.innerHTML = `<div class="modal prog-dlg" role="dialog" aria-modal="true" aria-label="${esc(title)}" style="width:440px">
    <header>${esc(title)}</header>
    <div class="body"><div class="prog-text">${esc(text)}</div><div class="prog-bar"><i></i></div><div class="prog-sub"></div></div>
    <footer><button class="btn" data-cancel>Cancel</button></footer></div>`;
  const h = { cancelled: false, onCancel: null };
  const bar = back.querySelector('.prog-bar i'), txt = back.querySelector('.prog-text'), sub = back.querySelector('.prog-sub');
  const key = (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); cancel(); } };
  const cancel = () => { if (h.cancelled) return; h.cancelled = true; txt.textContent = 'Stopping…'; back.querySelector('[data-cancel]').disabled = true; h.onCancel?.(); };
  back.querySelector('[data-cancel]').addEventListener('click', cancel);
  document.addEventListener('keydown', key, true);
  document.body.appendChild(back);
  h.set = (frac, t, s) => {
    if (frac == null || frac < 0) bar.parentNode.classList.add('indet'); else { bar.parentNode.classList.remove('indet'); bar.style.width = Math.round(Math.min(1, frac) * 100) + '%'; }
    if (t != null && !h.cancelled) txt.textContent = t;
    if (s != null) sub.textContent = s;
  };
  h.close = () => { back.remove(); document.removeEventListener('keydown', key, true); };
  return h;
}

export class Cancelled extends Error { constructor() { super('Cancelled'); this.name = 'Cancelled'; } }
export const nextFrame = () => new Promise(r => setTimeout(r, 0));
