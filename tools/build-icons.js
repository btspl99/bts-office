// Builds app/common/icons.js from lucide-static (ISC licensed) icon set.
const fs=require('fs'),path=require('path');
const dir=path.join(__dirname,'..','node_modules','lucide-static','icons');
const map={};
for(const f of fs.readdirSync(dir)){ if(!f.endsWith('.svg'))continue;
  const s=fs.readFileSync(path.join(dir,f),'utf8');
  const inner=s.replace(/^[\s\S]*?<svg[^>]*>/,'').replace(/<\/svg>\s*$/,'').replace(/\s+/g,' ').trim();
  map[f.slice(0,-4)]=inner; }
fs.writeFileSync(path.join(__dirname,'..','app','common','icons.js'),
 '/* Icons: Lucide (ISC License) https://lucide.dev */\nwindow.BTS_ICONS='+JSON.stringify(map)+';\n'+
 'window.icon=function(n,s){s=s||18;var b=window.BTS_ICONS[n]||window.BTS_ICONS["square"];return \'<svg class="ic" width="\'+s+\'" height="\'+s+\'" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round">\'+b+\'</svg>\';};\n'+
 'document.addEventListener("DOMContentLoaded",function(){document.querySelectorAll("[data-icon]").forEach(function(e){e.insertAdjacentHTML("afterbegin",icon(e.dataset.icon,+e.dataset.iconSize||undefined));});});\n');
console.log(Object.keys(map).length,'icons');
