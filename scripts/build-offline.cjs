const fs=require('node:fs');
const path=require('node:path');
const root=path.resolve(__dirname,'..');
let html=fs.readFileSync(path.join(root,'index.html'),'utf8');
html=html.replace('<link rel="stylesheet" href="style.css">',()=>`<style>${fs.readFileSync(path.join(root,'style.css'),'utf8')}</style>`);
for(const file of ['jszip.min.js','core.js','excel.js','app.js']){
 const script=fs.readFileSync(path.join(root,file),'utf8').replace(/<\/script/gi,'<\\/script');
 html=html.replace(`<script src="${file}"></script>`,()=>`<script>${script}</script>`);
}
fs.writeFileSync(path.join(root,'volunteer-selection-offline.html'),html);
console.log(`Generated standalone HTML (${Buffer.byteLength(html)} bytes)`);
