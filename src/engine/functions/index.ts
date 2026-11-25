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

export function functionList() {
  return [...FUNCTIONS.values()].filter((f) => f.category !== 'Internal').sort((a, b) => a.name.localeCompare(b.name));
}
