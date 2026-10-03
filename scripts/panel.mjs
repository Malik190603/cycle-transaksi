// Menyusun halaman panel admin (Upload, Facility, Config) dari src/: tampilan yang sama dengan aplikasi,
// dijalankan dalam mode panel dan ditanam ke backend/Code.gs sebagai PANEL_HTML_ (lihat doGet di Main.gs).
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { minify } from 'terser';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SRC = path.join(ROOT, 'src');
const daftar = (dir, ext) => fs.readdirSync(path.join(SRC, dir)).filter((f) => f.endsWith(ext)).sort().map((f) => fs.readFileSync(path.join(SRC, dir, f), 'utf8'));

export async function panelHtml() {
  const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
  // huruf: file lokal tidak ikut ke Apps Script, jadi diambil dari Google Fonts
  const css = daftar('css', '.css').join('\n').replace(/@font-face\s*\{[^}]*\}\s*/g, '');
  const js = (await minify("(function () {\n'use strict';\n" + daftar('js', '.js').join('\n') + '\n})();', { compress: { passes: 1 }, mangle: true, format: { comments: false } })).code;
  const cfg = { version: pkg.version, panel: true, updateRepo: '', nativeBase: '', build: 0, serverUrl: '', demoSaja: false, xlsxUrl: 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/' + pkg.dependencies.xlsx + '/xlsx.full.min.js' };
  const body = fs.readFileSync(path.join(SRC, 'index.html'), 'utf8').replace(/^[\s\S]*<body>/, '').replace(/<\/body>[\s\S]*$/, '').replace(/\s*<script src="config\.js"><\/script>\s*<script src="app\.js"><\/script>\s*/, '\n');
  if (/<script/.test(body)) throw new Error('kerangka src/index.html berubah');
  const aman = (s) => s.replace(/<\/script/gi, '<\\/script').replace(/<!--/g, '<\\!--');
  return `<!DOCTYPE html><html lang="id"><head><meta charset="utf-8"><base target="_top">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap">
<script>try{var t=JSON.parse(localStorage.getItem('ct.tema'));if(t==='gelap')document.documentElement.setAttribute('data-tema','gelap');}catch(e){}</script>
<style>${css}
#app{max-width:640px}</style></head><body>${body.trim()}
<script>window.CT_CONFIG=${JSON.stringify(cfg)};</script>
<script>${aman(js)}</script></body></html>`;
}
