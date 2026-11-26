// Row/column outline groups: group, ungroup, collapse and expand.
// Collapsing a group hides its members through the sheet's ordinary hidden set,
// so every consumer of hidden rows/columns (layout, navigation, copy, charts)
// keeps working without special cases. Groups may nest.

import type { OutlineGroup, Sheet } from '../workbook';

export type OutlineAxis = 'row' | 'col';

const groupsOf = (sh: Sheet, axis: OutlineAxis) => (axis === 'row' ? sh.rowGroups : sh.colGroups);
const hiddenOf = (sh: Sheet, axis: OutlineAxis) => (axis === 'row' ? sh.hiddenRows : sh.hiddenCols);

/** Hard cap so grouping a whole column of 5M rows can't lock up the page. */
const MAX_SPAN = 200000;

/** Nesting depth of a group (1 = outermost). */
export function groupLevel(groups: OutlineGroup[], g: OutlineGroup): number {
  let level = 1;
  for (const o of groups) if (o !== g && o.s <= g.s && o.e >= g.e && (o.s !== g.s || o.e !== g.e)) level++;
  return level;
}

export function maxOutlineLevel(groups: OutlineGroup[]): number {
  let m = 0;
  for (const g of groups) m = Math.max(m, groupLevel(groups, g));
  return m;
}

/** Adds a group over [s, e]. Returns false when it already exists or would partially overlap another group. */
export function addGroup(sh: Sheet, axis: OutlineAxis, s: number, e: number): boolean {
  if (e < s) [s, e] = [e, s];
  e = Math.min(e, s + MAX_SPAN);
  const gs = groupsOf(sh, axis);
  for (const g of gs) {
    if (g.s === s && g.e === e) return false;
    const overlaps = g.s <= e && s <= g.e;
    const nested = (g.s <= s && g.e >= e) || (s <= g.s && e >= g.e);
    if (overlaps && !nested) return false;
  }
  gs.push({ s, e, collapsed: false });
  gs.sort((a, b) => a.s - b.s || b.e - a.e);
  return true;
}

/** Removes the groups inside [s, e] (or, if none, the innermost group containing it), expanding them first. */
export function removeGroups(sh: Sheet, axis: OutlineAxis, s: number, e: number): number {
  const gs = groupsOf(sh, axis);
  let victims = gs.filter((g) => g.s >= s && g.e <= e);
  if (!victims.length) {
    const containing = gs.filter((g) => g.s <= s && g.e >= e).sort((a, b) => a.e - a.s - (b.e - b.s));
    if (containing.length) victims = [containing[0]];
  }
  for (const v of victims) if (v.collapsed) setCollapsed(sh, axis, v, false);
  const keep = gs.filter((g) => !victims.includes(g));
  if (axis === 'row') sh.rowGroups = keep;
  else sh.colGroups = keep;
  return victims.length;
}

/** Collapses or expands one group, keeping members of other collapsed groups hidden. */
export function setCollapsed(sh: Sheet, axis: OutlineAxis, g: OutlineGroup, collapsed: boolean) {
  g.collapsed = collapsed;
  const hidden = hiddenOf(sh, axis);
  const gs = groupsOf(sh, axis);
  for (let i = g.s; i <= g.e; i++) {
    if (collapsed) hidden.add(i);
    else if (!gs.some((o) => o !== g && o.collapsed && i >= o.s && i <= o.e)) hidden.delete(i);
  }
}

/** The innermost group whose toggle sits at `index` (the row/column just after the group). */
export function groupAtToggle(sh: Sheet, axis: OutlineAxis, index: number): OutlineGroup | undefined {
  const cands = groupsOf(sh, axis).filter((g) => g.e === index - 1);
  if (!cands.length) return undefined;
  // a collapsed outer group hides inner toggles, so prefer the outermost collapsed one
  const collapsed = cands.filter((g) => g.collapsed).sort((a, b) => a.s - b.s);
  if (collapsed.length) return collapsed[0];
  return cands.sort((a, b) => b.s - a.s)[0];
}

export function toggleAt(sh: Sheet, axis: OutlineAxis, index: number): boolean {
  const g = groupAtToggle(sh, axis, index);
  if (!g) return false;
  setCollapsed(sh, axis, g, !g.collapsed);
  return true;
}

/** Shows outline levels up to `level`: deeper groups collapse, shallower ones expand. 0 expands everything. */
export function showLevel(sh: Sheet, axis: OutlineAxis, level: number) {
  const gs = groupsOf(sh, axis);
  const levels = gs.map((g) => groupLevel(gs, g));
  // expand first so the hidden set is rebuilt cleanly, then collapse
  gs.forEach((g) => g.collapsed && setCollapsed(sh, axis, g, false));
  if (level <= 0) return;
  gs.forEach((g, i) => levels[i] >= level && setCollapsed(sh, axis, g, true));
}
