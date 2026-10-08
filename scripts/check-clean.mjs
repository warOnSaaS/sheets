// Fails if any private name (from .names or instances/*.names, both git-ignored) appears in a tracked file.
import { execSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const dir = 'instances';
const lists = [...(fs.existsSync('.names') ? ['.names'] : []), ...(fs.existsSync(dir) ? fs.readdirSync(dir).filter((f) => f.endsWith('.names')).map((f) => path.join(dir, f)) : [])];
const names = lists.flatMap((f) => fs.readFileSync(f, 'utf8').split('\n').map((l) => l.trim()).filter((l) => l && !l.startsWith('#')));
if (!names.length) {
  console.log('check:clean: no .names or instances/*.names files, nothing to check');
  process.exit(0);
}
const files = execSync('git ls-files', { encoding: 'utf8' }).split('\n').filter(Boolean);
const hits = [];
for (const f of files) {
  if (!fs.existsSync(f) || fs.lstatSync(f).isSymbolicLink()) continue;
  const text = fs.readFileSync(f, 'utf8').toLowerCase();
  // Whole words only, so a name never matches inside a hash or a longer word.
  for (const n of names) if (new RegExp(`(^|[^a-z0-9])${n.toLowerCase().replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}($|[^a-z0-9])`).test(text)) hits.push(`${f}: "${n}"`);
}
if (hits.length) {
  console.error(`check:clean: instance-specific names in tracked files:\n  ${hits.join('\n  ')}`);
  process.exit(1);
}
console.log(`check:clean: ${files.length} files clean of ${names.length} instance names`);
