import React, { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { getStore, useApp } from '../store';
import { fnSuggestions, argHint, signatureParts, formulaRefs, canInsertRef, refBeforeCaret } from '../formulaAssist';
import { Style } from '../../engine/types';

interface Props {
  mode: 'cell' | 'bar';
  box?: { x: number; y: number; w: number; h: number };
  zoom?: number;
  style?: Style;
  onCommit: (move: 'down' | 'up' | 'right' | 'left' | 'none', all?: boolean) => void;
  onCancel: () => void;
  onPointMove?: (dir: 'up' | 'down' | 'left' | 'right', extend: boolean) => void;
  inputRef?: React.RefObject<HTMLTextAreaElement>;
}

let measure: CanvasRenderingContext2D | null = null;
function textWidth(t: string, font: string) {
  measure = measure ?? document.createElement('canvas').getContext('2d');
  measure!.font = font;
  return Math.max(...t.split('\n').map((l) => measure!.measureText(l).width));
}

/** Toggle $ anchoring for the reference under the caret: A1 → $A$1 → A$1 → $A1 → A1 */
function cycleAbsolute(text: string, caret: number): { text: string; caret: number } {
  const re = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    const s = m.index, e = s + m[0].length;
    if (caret >= s && caret <= e) {
      const [ca, ra] = [m[1] === '$', m[3] === '$'];
      const next = !ca && !ra ? ['$', '$'] : ca && ra ? ['', '$'] : !ca && ra ? ['$', ''] : ['', ''];
      const rep = next[0] + m[2] + next[1] + m[4];
      return { text: text.slice(0, s) + rep + text.slice(e), caret: s + rep.length };
    }
  }
  return { text, caret };
}

export function FormulaEditor({ mode, box, zoom = 1, style, onCommit, onCancel, onPointMove, inputRef }: Props) {
  const store = getStore();
  const edit = useApp((s) => s.edit)!;
  const ownRef = useRef<HTMLTextAreaElement>(null);
  const ref = inputRef ?? ownRef;
  const [acIndex, setAcIndex] = useState(0);
  const [acClosed, setAcClosed] = useState(false);
  const active = edit && edit.source === mode;
  const names = [...store.wb.names.values()].map((n) => n.name);
  const suggestions = useMemo(() => (edit && active && !acClosed ? fnSuggestions(edit.text, edit.caret, names) : null), [edit?.text, edit?.caret, active, acClosed]);
  const hint = useMemo(() => (edit && active ? argHint(edit.text, edit.caret) : null), [edit?.text, edit?.caret, active]);
  const refs = useMemo(() => (edit ? formulaRefs(edit.text) : []), [edit?.text]);

  useEffect(() => setAcIndex(0), [suggestions?.prefix]);
  useEffect(() => setAcClosed(false), [edit?.text]);

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el || !active) return;
    if (document.activeElement !== el) el.focus();
    if (el.selectionStart !== edit.caret || el.selectionEnd !== edit.caret) el.setSelectionRange(edit.caret, edit.caret);
  }, [edit?.caret, active, edit?.text]);

  if (!edit) return null;

  const fontSize = (style?.size ?? 13) * (mode === 'cell' ? zoom : 1);
  const font = `${style?.i ? 'italic ' : ''}${style?.b ? '600 ' : ''}${fontSize}px ${mode === 'bar' ? getComputedStyle(document.body).getPropertyValue('--mono') || 'monospace' : getComputedStyle(document.body).fontFamily}`;

  const setText = (text: string, caret: number) => store.set({ edit: { ...edit, text, caret, source: mode } });

  const accept = (i: number) => {
    const sug = suggestions!;
    const item = sug.items[i];
    const after = edit.text.slice(edit.caret);
    const insert = item.kind === 'fn' ? item.name + (after.startsWith('(') ? '' : '(') : item.name;
    const text = edit.text.slice(0, sug.start) + insert + after;
    setText(text, sug.start + insert.length);
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    // the editor owns all keys while editing; never let them reach the grid's handler
    e.stopPropagation();
    const k = e.key;
    const caret = e.currentTarget.selectionStart;
    if (suggestions && suggestions.items.length) {
      if (k === 'ArrowDown') { e.preventDefault(); setAcIndex((i) => (i + 1) % suggestions.items.length); return; }
      if (k === 'ArrowUp') { e.preventDefault(); setAcIndex((i) => (i - 1 + suggestions.items.length) % suggestions.items.length); return; }
      if (k === 'Tab' || (k === 'Enter' && !e.shiftKey && !e.altKey)) { e.preventDefault(); accept(acIndex); return; }
      if (k === 'Escape') { e.preventDefault(); setAcClosed(true); return; }
    }
    if (k === 'Enter') {
      if (e.altKey) {
        e.preventDefault();
        setText(edit.text.slice(0, caret) + '\n' + edit.text.slice(e.currentTarget.selectionEnd), caret + 1);
        return;
      }
      e.preventDefault();
      onCommit(e.shiftKey ? 'up' : 'down', e.ctrlKey || e.metaKey);
      return;
    }
    if (k === 'Tab') {
      e.preventDefault();
      onCommit(e.shiftKey ? 'left' : 'right');
      return;
    }
    if (k === 'Escape') {
      e.preventDefault();
      onCancel();
      return;
    }
    if (k === 'F4') {
      e.preventDefault();
      const r = cycleAbsolute(edit.text, caret);
      setText(r.text, r.caret);
      return;
    }
    if (k.startsWith('Arrow') && mode === 'cell' && edit.enterMode && !e.altKey) {
      const dir = k.slice(5).toLowerCase() as 'up' | 'down' | 'left' | 'right';
      if (edit.text.startsWith('=') && (canInsertRef(edit.text, caret) || refBeforeCaret(edit.text, caret)) && onPointMove) {
        e.preventDefault();
        onPointMove(dir, e.shiftKey);
        return;
      }
      if (!edit.text.startsWith('=')) {
        e.preventDefault();
        onCommit(dir);
        return;
      }
    }
    if ((e.metaKey || e.ctrlKey) && k.toLowerCase() === 'z') {
      // let the textarea handle its own undo
      e.stopPropagation();
    }
  };

  const onChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(e.target.value, e.target.selectionStart);
  };

  const onSelect = (e: React.SyntheticEvent<HTMLTextAreaElement>) => {
    const c = e.currentTarget.selectionStart;
    if (c !== edit.caret && active) store.set({ edit: { ...edit, caret: c, source: mode, enterMode: false } });
  };

  // coloured mirror text (references in the same colours as the grid highlights)
  const pieces: React.ReactNode[] = [];
  if (edit.text.startsWith('=')) {
    let pos = 0;
    refs.forEach((r, i) => {
      if (r.start > pos) pieces.push(<span key={'t' + i}>{edit.text.slice(pos, r.start)}</span>);
      pieces.push(<span key={'r' + i} style={{ color: r.color, fontWeight: 600 }}>{edit.text.slice(r.start, r.end)}</span>);
      pos = r.end;
    });
    pieces.push(<span key="end">{edit.text.slice(pos)}</span>);
  }
  const colored = edit.text.startsWith('=');

  let width: number | undefined, height: number | undefined;
  if (mode === 'cell' && box) {
    width = Math.min(Math.max(box.w, textWidth(edit.text + 'MM', font) + 14), 720);
    const lines = edit.text.split('\n').length;
    height = Math.max(box.h, lines * fontSize * 1.4 + 8);
  }

  const common: React.CSSProperties = {
    font, lineHeight: mode === 'cell' ? `${fontSize * 1.38}px` : '18px', letterSpacing: 'normal',
  };

  const acPos: React.CSSProperties = mode === 'cell' && box ? { left: box.x, top: box.y + (height ?? box.h) + 4 } : { left: 160, top: 38 };
  const hintPos: React.CSSProperties = mode === 'cell' && box ? { left: box.x, top: box.y - 58 < 0 ? box.y + (height ?? box.h) + 4 : box.y - 58 } : { left: 160, top: 40 };

  return (
    <>
      {mode === 'cell' && box ? (
        <div style={{ position: 'absolute', left: box.x, top: box.y, width, height, zIndex: 15 }}>
          {colored && (
            <div aria-hidden className="cell-editor" style={{ ...common, position: 'absolute', inset: 0, color: 'var(--grid-text)', pointerEvents: 'none', zIndex: 0 }}>
              {pieces}
            </div>
          )}
          <textarea
            ref={ref}
            className="cell-editor"
            spellCheck={false}
            value={edit.text}
            onChange={onChange}
            onKeyDown={onKeyDown}
            onSelect={onSelect}
            onMouseDown={(e) => e.stopPropagation()}
            style={{
              ...common, position: 'absolute', inset: 0, width: '100%', height: '100%',
              color: colored ? 'transparent' : 'var(--grid-text)', caretColor: 'var(--grid-text)',
              background: colored ? 'transparent' : 'var(--grid-bg)', fontWeight: style?.b ? 600 : undefined, fontStyle: style?.i ? 'italic' : undefined,
              textAlign: style?.h ?? 'left',
            }}
          />
        </div>
      ) : (
        <div style={{ position: 'relative', flex: 1, display: 'flex' }}>
          {colored && active && (
            <div aria-hidden className="fbar-input" style={{ position: 'absolute', inset: 0, pointerEvents: 'none', color: 'var(--text)' }}>{pieces}</div>
          )}
          <textarea
            ref={ref}
            className="fbar-input"
            spellCheck={false}
            value={edit.text}
            onChange={onChange}
            onKeyDown={onKeyDown}
            onSelect={onSelect}
            style={{ color: colored && active ? 'transparent' : undefined, caretColor: 'var(--text)', position: 'relative' }}
          />
        </div>
      )}
      {active && suggestions && suggestions.items.length > 0 && (
        <div className="popover ac-list" style={{ ...acPos, position: 'absolute' }} onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}>
          {suggestions.items.map((it, i) => (
            <div key={it.name} className={`ac-item ${i === acIndex ? 'hl' : ''}`} onMouseEnter={() => setAcIndex(i)} onClick={() => accept(i)}>
              <b>{it.name}{it.kind === 'fn' ? '(' : ''}</b>
              <span>{it.desc}</span>
            </div>
          ))}
        </div>
      )}
      {active && hint && !(suggestions && suggestions.items.length) && (
        <div className="arg-hint" style={hintPos}>
          <div>
            <b>{hint.def.name}</b>(
            {signatureParts(hint.def).map((p, i, arr) => {
              const repeatIdx = hint.def.args.findIndex((a) => a.repeat);
              const cur = i === hint.arg || (repeatIdx >= 0 && i === repeatIdx && hint.arg >= repeatIdx) || (i === arr.length - 1 && hint.arg > i && hint.def.args[i]?.repeat);
              return (
                <span key={i}>
                  <span className={cur ? 'cur' : ''}>{p}</span>
                  {i < arr.length - 1 ? ', ' : ''}
                </span>
              );
            })}
            )
          </div>
          <div className="desc">
            {hint.def.args[Math.min(hint.arg, hint.def.args.length - 1)] ? <><b>{hint.def.args[Math.min(hint.arg, hint.def.args.length - 1)].name}</b>: {hint.def.args[Math.min(hint.arg, hint.def.args.length - 1)].desc}</> : hint.def.desc}
          </div>
        </div>
      )}
    </>
  );
}
