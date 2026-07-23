const fs = require('fs');
const { JSDOM } = require('jsdom');

const html = fs.readFileSync('/mnt/user-data/outputs/index.html', 'utf8');

// Tách phần JS trong script cuối (không chạy React CDN, ta tự cấp React)
const scriptMatch = html.match(/<script>\n\(function\(\)\{([\s\S]*)\}\)\(\);\n<\/script>/);
if (!scriptMatch) { console.log('NO SCRIPT MATCH'); process.exit(1); }
let inner = scriptMatch[1];

// Cấp React/ReactDOM/XLSX từ node_modules, và localStorage + document từ jsdom
const dom = new JSDOM('<!DOCTYPE html><div id="root"></div>', { url: 'http://localhost' });
global.window = dom.window;
global.document = dom.window.document;
global.navigator = dom.window.navigator;
const store = {};
global.localStorage = { getItem: k => (k in store ? store[k] : null), setItem: (k,v)=>{store[k]=String(v)}, removeItem: k=>{delete store[k]} };
dom.window.localStorage = global.localStorage;

const React = require('react');
const ReactDOMServer = require('react-dom/server');
global.window.React = React;
// tạo ReactDOM.createRoot giả để bắt component
let captured = null;
global.window.ReactDOM = { createRoot: () => ({ render: (el) => { captured = el; } }) };
global.window.XLSX = {};

// chạy inner trong 1 function scope
try {
  const fn = new Function('React','ReactDOM','XLSX','window','document','navigator','localStorage', inner
    .replace('var React = window.React;','')
    .replace('var ReactDOM = window.ReactDOM;','')
    .replace('var XLSX = window.XLSX;','')
  );
  fn(React, global.window.ReactDOM, global.window.XLSX, global.window, global.document, global.navigator, global.localStorage);
  if (!captured) { console.log('No component captured'); process.exit(1); }
  const out = ReactDOMServer.renderToStaticMarkup(captured);
  console.log('RENDER OK len', out.length, 'undefined?', out.includes('undefined'), 'hasSidebar?', out.includes('TRẠM ĐIỀU PHỐI'), 'hasSVG?', out.includes('<svg'));
} catch(e) {
  console.log('ERROR:', e.message);
  console.log(e.stack.split('\n').slice(0,4).join('\n'));
}
