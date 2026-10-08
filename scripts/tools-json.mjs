// Writes tools.json (the suite's catalogue format, packages/tools) from lib/tools.mjs. --check fails if stale.
import fs from 'node:fs';
import { listTools } from '../lib/tools.mjs';

const tests = { 'sheets.sync': 'test/server.test.mjs', 'sheets.set_presence': 'test/server.test.mjs' };
const doc = {
  $schema: 'https://raw.githubusercontent.com/warOnSaaS/suite/main/packages/tools/tools.schema.json',
  app: 'sheets',
  version: 1,
  tools: listTools().map((t) => ({ name: t.name, title: t.title, description: t.description, input: t.inputJson, output: t.outputJson, scope: t.scope, confirm: t.confirm, emits: t.emits, test: tests[t.name] ?? 'test/tools.test.mjs' })),
};
const text = `${JSON.stringify(doc, null, 2)}\n`;
if (process.argv.includes('--check')) {
  if ((fs.existsSync('tools.json') ? fs.readFileSync('tools.json', 'utf8') : '') !== text) { console.error('tools.json is out of date: run npm run tools:json'); process.exit(1); }
  console.log(`tools.json is current (${doc.tools.length} tools)`);
} else { fs.writeFileSync('tools.json', text); console.log(`wrote tools.json (${doc.tools.length} tools)`); }
process.exit(0);
