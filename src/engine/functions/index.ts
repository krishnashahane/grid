import './math';
import './logic';
import './lookup';
import './text';
import './stats';
import './date';
import './array';
import './finance';
import { FUNCTIONS } from '../evaluate';

export { FUNCTIONS };

// Criteria arguments of the *IF/*IFS family broadcast over arrays, e.g. SUMIFS(C:C, A:A, UNIQUE(A:A)).
const criteriaLift: Record<string, (i: number) => boolean> = {
  SUMIF: (i) => i === 1,
  COUNTIF: (i) => i === 1,
  AVERAGEIF: (i) => i === 1,
  SUMIFS: (i) => i >= 2 && i % 2 === 0,
  AVERAGEIFS: (i) => i >= 2 && i % 2 === 0,
  MAXIFS: (i) => i >= 2 && i % 2 === 0,
  MINIFS: (i) => i >= 2 && i % 2 === 0,
  COUNTIFS: (i) => i % 2 === 1,
};
for (const [name, lift] of Object.entries(criteriaLift)) {
  const d = FUNCTIONS.get(name);
  if (d) d.lift = lift;
}

export function functionList() {
  return [...FUNCTIONS.values()].filter((f) => f.category !== 'Internal').sort((a, b) => a.name.localeCompare(b.name));
}
