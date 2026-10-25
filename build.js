// 把程式庫、規則、畫面合成單一 index.html（離線可用，雙擊即開）
const fs = require('fs');
const rd = (p) => fs.readFileSync(p, 'utf8').replace(/<\/script/gi, '<\\/script');
let html = fs.readFileSync('src/template.html', 'utf8');
const libs = ['node_modules/xlsx/dist/xlsx.full.min.js', 'node_modules/exceljs/dist/exceljs.min.js']
  .map((p) => '<script>' + rd(p) + '</script>').join('\n');
html = html.replace('<!--LIBS-->', () => libs)
  .replace('/*CORE*/', () => rd('src/core.js'))
  .replace('/*APP*/', () => rd('src/app.js'));
fs.writeFileSync('index.html', html);
console.log('index.html', (html.length / 1024 / 1024).toFixed(2) + ' MB');
