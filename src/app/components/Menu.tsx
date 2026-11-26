import React, { useEffect, useLayoutEffect, useRef, useState } from 'react';

export interface MenuItem {
  label?: string;
  icon?: React.ReactNode;
  kbd?: string;
  onClick?: () => void;
  danger?: boolean;
  disabled?: boolean;
  sep?: boolean;
  header?: string;
}

export function Popover({ x, y, onClose, children, className = '', anchorRight = false }: { x: number; y: number; onClose: () => void; children: React.ReactNode; className?: string; anchorRight?: boolean }) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ x, y });
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    let nx = anchorRight ? x - r.width : x, ny = y;
    if (nx + r.width > window.innerWidth - 8) nx = window.innerWidth - r.width - 8;
    if (ny + r.height > window.innerHeight - 8) ny = Math.max(8, y - r.height);
    setPos({ x: Math.max(8, nx), y: ny });
  }, [x, y, anchorRight]);
  useEffect(() => {
    const down = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        onClose();
      }
    };
    setTimeout(() => window.addEventListener('mousedown', down), 0);
    window.addEventListener('keydown', key, true);
    return () => {
      window.removeEventListener('mousedown', down);
      window.removeEventListener('keydown', key, true);
    };
  }, [onClose]);
  return (
    <div
      ref={ref}
      className={`popover ${className}`}
      style={{ left: pos.x, top: pos.y }}
      onContextMenu={(e) => { e.preventDefault(); e.stopPropagation(); }}
      onMouseDown={(e) => e.stopPropagation()}
      onDoubleClick={(e) => e.stopPropagation()}
      onKeyDown={(e) => { if (e.key !== 'Escape') e.stopPropagation(); }}
    >
      {children}
    </div>
  );
}

export function Menu({ x, y, items, onClose, anchorRight }: { x: number; y: number; items: MenuItem[]; onClose: () => void; anchorRight?: boolean }) {
  return (
    <Popover x={x} y={y} onClose={onClose} className="menu" anchorRight={anchorRight}>
      {items.map((it, i) =>
        it.sep ? (
          <div key={i} className="menu-sep" />
        ) : it.header ? (
          <div key={i} className="menu-label">{it.header}</div>
        ) : (
          <button
            key={i}
            className={`menu-item ${it.danger ? 'danger' : ''}`}
            disabled={it.disabled}
            style={it.disabled ? { opacity: 0.45, cursor: 'default' } : undefined}
            onClick={() => {
              if (it.disabled) return;
              onClose();
              it.onClick?.();
            }}
          >
            {it.icon}
            <span>{it.label}</span>
            {it.kbd && <span className="kbd">{it.kbd}</span>}
          </button>
        ),
      )}
    </Popover>
  );
}

export function useMenu() {
  const [menu, setMenu] = useState<{ x: number; y: number; items: MenuItem[]; anchorRight?: boolean } | null>(null);
  const el = menu ? <Menu {...menu} onClose={() => setMenu(null)} /> : null;
  return { open: (x: number, y: number, items: MenuItem[], anchorRight = false) => setMenu({ x, y, items, anchorRight }), el, close: () => setMenu(null) };
}

export const isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);
export const MOD = isMac ? '⌘' : 'Ctrl+';
