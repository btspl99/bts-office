/* BTS Sheets — extra Excel functions missing from HyperFormula 2.7 (XLOOKUP, TEXTJOIN, RANK, AVERAGEIFS, VALUE …) */
(function () {
  const HF = window.HyperFormula; if (!HF) return;
  const E = (t) => new HF.CellError(HF.ErrorType[t]);
  const raw = (v) => (v && typeof v === 'object' && 'val' in v) ? v.val : v;      // rich numbers (dates, %, currency)
  const isEmpty = (v) => v === undefined || v === null || typeof v === 'symbol' || v === '';
  const isErr = (v) => v instanceof HF.CellError;
  const flat = (a) => (a && a.valuesFromTopLeftCorner) ? a.valuesFromTopLeftCorner().map(raw) : [raw(a)];
  const rows = (a) => (a && a.data) ? a.data.map(r => r.map(raw)) : [[raw(a)]];
  const nums = (arr) => arr.filter(v => typeof v === 'number');

  function criterion(c) {
    c = raw(c);
    if (typeof c === 'number' || typeof c === 'boolean') return (v) => raw(v) === c;
    const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(String(c)); const op = m[1] || '=', rhs = m[2];
    const n = rhs !== '' && !isNaN(+rhs) ? +rhs : null;
    let re = null; if (n === null && /[*?]/.test(rhs)) re = new RegExp('^' + rhs.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*') + '$', 'i');
    return (v) => {
      v = raw(v);
      if (n !== null) {
        if (typeof v !== 'number') return op === '<>';
        switch (op) { case '=': return v === n; case '<>': return v !== n; case '<': return v < n; case '>': return v > n; case '<=': return v <= n; case '>=': return v >= n; }
      }
      const s = isEmpty(v) ? '' : String(v);
      if (op === '=' ) return re ? re.test(s) : (rhs === '' ? isEmpty(v) : s.toLowerCase() === rhs.toLowerCase());
      if (op === '<>') return re ? !re.test(s) : (rhs === '' ? !isEmpty(v) : s.toLowerCase() !== rhs.toLowerCase());
      if (typeof v !== 'string') return false;
      const cmp = s.localeCompare(rhs, undefined, { sensitivity: 'base' });
      return op === '<' ? cmp < 0 : op === '>' ? cmp > 0 : op === '<=' ? cmp <= 0 : cmp >= 0;
    };
  }
  function percentileInc(arr, k) { arr = nums(arr).sort((a, b) => a - b); if (!arr.length || k < 0 || k > 1) return E('NUM'); const p = (arr.length - 1) * k, lo = Math.floor(p), hi = Math.ceil(p); return arr[lo] + (arr[hi] - arr[lo]) * (p - lo); }
  function linreg(ys, xs) {
    const Y = flat(ys), X = flat(xs); if (Y.length !== X.length) return null;
    const px = [], py = []; for (let i = 0; i < Y.length; i++) if (typeof Y[i] === 'number' && typeof X[i] === 'number') { px.push(X[i]); py.push(Y[i]); }
    const n = px.length; if (n < 2) return null; const mx = px.reduce((a, b) => a + b, 0) / n, my = py.reduce((a, b) => a + b, 0) / n;
    let sxy = 0, sxx = 0; for (let i = 0; i < n; i++) { sxy += (px[i] - mx) * (py[i] - my); sxx += (px[i] - mx) ** 2; }
    if (!sxx) return null; const b = sxy / sxx; return { b, a: my - b * mx };
  }
  function valueOf(s) {
    s = raw(s); if (typeof s === 'number') return s; if (typeof s === 'boolean') return E('VALUE');
    if (isEmpty(s)) return 0;
    const p = window.SU ? SU.parseInput(String(s)) : { value: +s };
    return typeof p.value === 'number' ? p.value : E('VALUE');
  }
  function lookupIn(arr, v, mode) {
    v = raw(v); const eq = (a) => typeof a === typeof v && (typeof a === 'string' ? a.toLowerCase() === v.toLowerCase() : a === v);
    if (mode === 2) { const f = criterion(v); return arr.findIndex(a => f(a)); }
    let i = arr.findIndex(eq); if (i >= 0 || mode === 0) return i;
    let best = -1;
    for (let j = 0; j < arr.length; j++) { const a = arr[j]; if (typeof a !== typeof v) continue; const c = typeof a === 'string' ? a.localeCompare(v) : a - v; if (mode === -1 && c < 0 && (best < 0 || (typeof a === 'string' ? a.localeCompare(arr[best]) > 0 : a > arr[best]))) best = j; if (mode === 1 && c > 0 && (best < 0 || (typeof a === 'string' ? a.localeCompare(arr[best]) < 0 : a < arr[best]))) best = j; }
    return best;
  }

  class BTSPlugin extends HF.FunctionPlugin {
    rankEq(ast, st) { return this.runFunction(ast.args, st, this.metadata('RANK.EQ'), (n, range, order) => { const v = nums(flat(range)); if (!v.includes(n)) return E('NA'); let r = 1; for (const x of v) if (order ? x < n : x > n) r++; return r; }); }
    rankAvg(ast, st) { return this.runFunction(ast.args, st, this.metadata('RANK.AVG'), (n, range, order) => { const v = nums(flat(range)); const c = v.filter(x => x === n).length; if (!c) return E('NA'); let r = 1; for (const x of v) if (order ? x < n : x > n) r++; return r + (c - 1) / 2; }); }
    textjoin(ast, st) {
      return this.runFunction(ast.args, st, this.metadata('TEXTJOIN'), (d, ign, ...args) => {
        const out = []; for (const a of args) for (const v of flat(a)) { if (isErr(v)) return v; if (ign && isEmpty(v)) continue; out.push(isEmpty(v) ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v)); }
        return out.join(d);
      });
    }
    concat(ast, st) { return this.runFunction(ast.args, st, this.metadata('CONCAT'), (...args) => { let s = ''; for (const a of args) for (const v of flat(a)) { if (isErr(v)) return v; s += isEmpty(v) ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v); } return s; }); }
    averageifs(ast, st) {
      return this.runFunction(ast.args, st, this.metadata('AVERAGEIFS'), (avg, ...pairs) => {
        const vals = flat(avg); const tests = []; for (let i = 0; i + 1 < pairs.length; i += 2) tests.push([flat(pairs[i]), criterion(pairs[i + 1])]);
        let s = 0, n = 0; for (let i = 0; i < vals.length; i++) { if (typeof vals[i] !== 'number') continue; if (tests.every(([r, f]) => f(r[i]))) { s += vals[i]; n++; } }
        return n ? s / n : E('DIV_BY_ZERO');
      });
    }
    modeSngl(ast, st) { return this.runFunction(ast.args, st, this.metadata('MODE.SNGL'), (...args) => { const m = new Map(); let best = null, bc = 1; for (const a of args) for (const v of nums(flat(a))) { const c = (m.get(v) || 0) + 1; m.set(v, c); if (c > bc) { bc = c; best = v; } } return best === null ? E('NA') : best; }); }
    percentile(ast, st) { return this.runFunction(ast.args, st, this.metadata('PERCENTILE.INC'), (a, k) => percentileInc(flat(a), k)); }
    quartile(ast, st) { return this.runFunction(ast.args, st, this.metadata('QUARTILE.INC'), (a, q) => (q < 0 || q > 4) ? E('NUM') : percentileInc(flat(a), Math.floor(q) / 4)); }
    intercept(ast, st) { return this.runFunction(ast.args, st, this.metadata('INTERCEPT'), (ys, xs) => { const r = linreg(ys, xs); return r ? r.a : E('DIV_BY_ZERO'); }); }
    forecast(ast, st) { return this.runFunction(ast.args, st, this.metadata('FORECAST.LINEAR'), (x, ys, xs) => { const r = linreg(ys, xs); return r ? r.a + r.b * x : E('DIV_BY_ZERO'); }); }
    value(ast, st) { return this.runFunction(ast.args, st, this.metadata('VALUE'), (s) => valueOf(s)); }
    nfn(ast, st) { return this.runFunction(ast.args, st, this.metadata('N'), (v) => { v = raw(v); return typeof v === 'number' ? v : v === true ? 1 : 0; }); }
    xlookup(ast, st) {
      return this.runFunction(ast.args, st, this.metadata('XLOOKUP'), (v, look, ret, nf, mode, search) => {
        const L = flat(look), R = rows(ret); const vertical = (look.height ? look.height() : 1) > 1 || (look.width ? look.width() : 1) === 1;
        let arr = L, idx;
        if (search === -1) { idx = lookupIn(arr.slice().reverse(), v, mode); if (idx >= 0) idx = arr.length - 1 - idx; }
        else idx = lookupIn(arr, v, mode);
        if (idx < 0) return nf === undefined || nf === null || isEmpty(nf) ? E('NA') : nf;
        if (vertical) { if (R.length !== L.length) return E('VALUE'); const row = R[idx]; return row.length === 1 ? (isEmpty(row[0]) ? 0 : row[0]) : HF.SimpleRangeValue.onlyValues([row]); }
        if (R[0].length !== L.length) return E('VALUE');
        const col = R.map(r => r[idx]); return col.length === 1 ? (isEmpty(col[0]) ? 0 : col[0]) : HF.SimpleRangeValue.onlyValues(col.map(x => [x]));
      });
    }
    xlookupSize(ast, st) { return HF.ArraySize ? HF.ArraySize.scalar() : undefined; }
    lookup(ast, st) {
      return this.runFunction(ast.args, st, this.metadata('LOOKUP'), (v, look, res) => {
        const L = flat(look); let idx = lookupIn(L, v, -1); const exact = lookupIn(L, v, 0); if (exact >= 0) idx = exact;
        if (idx < 0) return E('NA'); const R = res ? flat(res) : L; const out = R[idx]; return isEmpty(out) ? 0 : out;
      });
    }
  }
  const P = (t, extra = {}) => Object.assign({ argumentType: t }, extra);
  BTSPlugin.implementedFunctions = {
    'RANK.EQ': { method: 'rankEq', parameters: [P('NUMBER'), P('RANGE'), P('NUMBER', { defaultValue: 0 })] },
    'RANK': { method: 'rankEq', parameters: [P('NUMBER'), P('RANGE'), P('NUMBER', { defaultValue: 0 })] },
    'RANK.AVG': { method: 'rankAvg', parameters: [P('NUMBER'), P('RANGE'), P('NUMBER', { defaultValue: 0 })] },
    'TEXTJOIN': { method: 'textjoin', parameters: [P('STRING'), P('BOOLEAN'), P('ANY')], repeatLastArgs: 1 },
    'CONCAT': { method: 'concat', parameters: [P('ANY')], repeatLastArgs: 1 },
    'AVERAGEIFS': { method: 'averageifs', parameters: [P('RANGE'), P('RANGE'), P('SCALAR')], repeatLastArgs: 2 },
    'MODE.SNGL': { method: 'modeSngl', parameters: [P('ANY')], repeatLastArgs: 1 },
    'MODE': { method: 'modeSngl', parameters: [P('ANY')], repeatLastArgs: 1 },
    'PERCENTILE.INC': { method: 'percentile', parameters: [P('RANGE'), P('NUMBER')] },
    'PERCENTILE': { method: 'percentile', parameters: [P('RANGE'), P('NUMBER')] },
    'QUARTILE.INC': { method: 'quartile', parameters: [P('RANGE'), P('NUMBER')] },
    'QUARTILE': { method: 'quartile', parameters: [P('RANGE'), P('NUMBER')] },
    'INTERCEPT': { method: 'intercept', parameters: [P('RANGE'), P('RANGE')] },
    'FORECAST.LINEAR': { method: 'forecast', parameters: [P('NUMBER'), P('RANGE'), P('RANGE')] },
    'FORECAST': { method: 'forecast', parameters: [P('NUMBER'), P('RANGE'), P('RANGE')] },
    'VALUE': { method: 'value', parameters: [P('SCALAR')] },
    'N': { method: 'nfn', parameters: [P('SCALAR')] },
    'XLOOKUP': { method: 'xlookup', parameters: [P('SCALAR'), P('RANGE'), P('RANGE'), P('SCALAR', { optionalArg: true }), P('NUMBER', { defaultValue: 0 }), P('NUMBER', { defaultValue: 1 })] },
    'LOOKUP': { method: 'lookup', parameters: [P('SCALAR'), P('RANGE'), P('RANGE', { optionalArg: true })] },
  };
  const names = {}; Object.keys(BTSPlugin.implementedFunctions).forEach(n => names[n] = n);
  try { HF.registerFunctionPlugin(BTSPlugin, { enGB: names, enUS: names }); } catch (e) { console.warn('BTS functions plugin:', e.message); }
})();
