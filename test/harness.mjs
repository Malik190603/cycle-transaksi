// Backend Apps Script yang asli, dijalankan di Node (layanan Google ditiru test/gas-mock.mjs).
// Dipakai uji tampilan (test/ui.mjs) dan pemeriksaan data contoh.
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
import { createGas } from './gas-mock.mjs';
import { seedDemo, DEMO } from '../demo/seed.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

export function backendSource(single) {
  const SRC = path.join(ROOT, 'backend', 'src');
  return single ? fs.readFileSync(path.join(ROOT, 'backend', 'Code.gs'), 'utf8')
    : fs.readdirSync(SRC).filter((f) => f.endsWith('.gs')).sort().map((f) => fs.readFileSync(path.join(SRC, f), 'utf8')).join('\n');
}

export function createBackend({ single = false, nama = DEMO.facility, kode = DEMO.kode, admin = DEMO.admin } = {}) {
  const code = backendSource(single).replace(/'ISI NAMA FACILITY'/, JSON.stringify(nama)).replace(/'ISI-KODE'/, JSON.stringify(kode)).replace(/'ISI\.NIK\.ADMIN'/, JSON.stringify(admin));
  const gas = createGas();
  const ctx = vm.createContext({ ...gas.globals });
  vm.runInContext(code, ctx, { filename: 'backend.js' });
  const post = (body) => JSON.parse(ctx.doPost({ postData: { contents: typeof body === 'string' ? body : JSON.stringify(body) } }).getContent());
  const raw = (action, ...args) => post({ action, args });
  const api = (action, ...args) => { const j = raw(action, ...args); if (!j.ok) throw new Error(action + ': ' + j.error); return j.result; };
  const fn = (name, ...args) => ctx[name](...args);
  const sheet = (n) => gas.active.getSheetByName(n);
  const rows = (n) => { const s = sheet(n); const lr = s.getLastRow(); return lr < 2 ? [] : s.getRange(2, 1, lr - 1, Math.max(1, s.getLastColumn())).getValues(); };
  return { gas, ctx, post, raw, api, fn, sheet, rows, worker: () => ctx.workerSemua() };
}

export function createDemoBackend(opts) {
  const be = createBackend();
  const info = seedDemo({ api: be.api, fn: be.fn, sheet: be.sheet }, opts);
  return Object.assign(be, { info });
}
