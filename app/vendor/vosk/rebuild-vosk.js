// Build helper (not loaded by the app). Regenerate app/vendor/vosk/vosk.js after upgrading vosk-browser:
//   node app/vendor/vosk/rebuild-vosk.js node_modules/vosk-browser/dist/vosk.js app/vendor/vosk/vosk.js
// Copies vosk-browser dist/vosk.js and patches its embedded worker for BTS Office:
//  1. keep the model in memory (MEMFS) instead of persisting a copy in IndexedDB (IDBFS) — BTS stores models on disk itself;
//  2. replace embind's code generation (new Function) with closures, so the page CSP does not need 'unsafe-eval'.
const fs = require('fs');
const [src, dst] = process.argv.slice(2);
let s = fs.readFileSync(src, 'utf8');
const m = s.match(/createBase64WorkerFactory\('([A-Za-z0-9+/=]+)'/);
let w = Buffer.from(m[1], 'base64').toString('utf8');

function swap(re, rep, label) { if (!re.test(w)) throw new Error('patch point not found: ' + label); w = w.replace(re, rep); }
swap(/this\.Vosk\.FS\.mount\(this\.Vosk\.IDBFS, \{\}, storagePath\);\s*return this\.Vosk\.syncFilesystem\(true\);/, '/* BTS: no IndexedDB persistence */', 'idbfs mount');
swap(/return this\.Vosk\.syncFilesystem\(false\);/, 'return; /* BTS: no IndexedDB persistence */', 'idbfs sync');

function swapRange(startStr, endStr, rep) {
  const i = w.indexOf(startStr); if (i < 0) throw new Error('patch point not found: ' + startStr);
  const j = w.indexOf(endStr, i); if (j < 0) throw new Error('patch end not found: ' + endStr);
  w = w.slice(0, i) + rep + w.slice(j);
}
swapRange('function createNamedFunction(name,body){', 'var emval_free_list=',
  'function createNamedFunction(name,body){name=makeLegalFunctionName(name);var f=function(){"use strict";return body.apply(this,arguments)};try{Object.defineProperty(f,"name",{value:name})}catch(e){}return f}');
swapRange('function craftInvokerFunction(', 'function heap32VectorToArray(',
  'function craftInvokerFunction(humanName,argTypes,classType,cppInvokerFunc,cppTargetFunc){var argCount=argTypes.length;if(argCount<2){throwBindingError("argTypes array size mismatch! Must at least get return value and \'this\' types!");}' +
  'var isClassMethodFunc=argTypes[1]!==null&&classType!==null;var needsDestructorStack=false;for(var i=1;i<argTypes.length;++i){if(argTypes[i]!==null&&argTypes[i].destructorFunction===undefined){needsDestructorStack=true;break}}' +
  'var returns=argTypes[0].name!=="void";var expected=argCount-2;' +
  'var f=function(){if(arguments.length!==expected){throwBindingError("function "+humanName+" called with "+arguments.length+" arguments, expected "+expected+" args!");}' +
  'var destructors=needsDestructorStack?[]:null;var wired=[cppTargetFunc];var thisWired;if(isClassMethodFunc){thisWired=argTypes[1].toWireType(destructors,this);wired.push(thisWired);}' +
  'var argsWired=new Array(expected);for(var i=0;i<expected;++i){argsWired[i]=argTypes[i+2].toWireType(destructors,arguments[i]);wired.push(argsWired[i]);}' +
  'var rv=cppInvokerFunc.apply(null,wired);' +
  'if(needsDestructorStack){runDestructors(destructors);}else{for(var i=isClassMethodFunc?1:2;i<argTypes.length;++i){var p=i===1?thisWired:argsWired[i-2];if(argTypes[i].destructorFunction!==null){argTypes[i].destructorFunction(p);}}}' +
  'if(returns){return argTypes[0].fromWireType(rv);}};try{Object.defineProperty(f,"name",{value:makeLegalFunctionName(humanName)})}catch(e){}return f}');
swapRange('function __emval_get_method_caller(', 'function __emval_incref(',
  'function __emval_get_method_caller(argCount,argTypes){var types=__emval_lookupTypes(argCount,argTypes);var retType=types[0];' +
  'var f=function(handle,name,destructors,args){var offset=0;var a=new Array(argCount-1);for(var i=0;i<argCount-1;++i){a[i]=types[1+i].readValueFromPointer(args+offset);offset+=types[i+1]["argPackAdvance"];}' +
  'var rv=handle[name].apply(handle,a);for(var i=0;i<argCount-1;++i){if(types[i+1]["deleteObject"]){types[i+1].deleteObject(a[i]);}}if(!retType.isVoid){return retType.toWireType(destructors,rv);}};return __emval_addMethodCaller(f)}');
if (/new Function|new_\(Function/.test(w)) throw new Error('eval left in worker');

s = s.replace(m[1], Buffer.from(w, 'utf8').toString('base64'));
fs.writeFileSync(dst, '/* vosk-browser 0.0.8 (Apache-2.0, https://github.com/ccoreilly/vosk-browser). Patched for BTS Office: the embedded worker keeps the model in memory only (no IndexedDB copy, BTS keeps models in userData/voice-models), and embind uses closures instead of new Function (no \'unsafe-eval\' needed). */\n' + s);
console.log('ok', s.length);
