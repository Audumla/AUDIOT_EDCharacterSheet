var GSUtils = (function () {

  /**
   * Fast string-equality lookup in a 2D array (e.g., from getDisplayValues()).
   * - col: 0-based column index OR header name (if headers=true).
   * - Returns the first match (row array by default), or index/both.
   */
  function findRowByColumn(data, col, needle, headers = false, ret = 'row') {
    if (!Array.isArray(data) || !Array.isArray(data[0])) throw new Error('data must be 2D');

    const colIndex = _resolveColIndexEq_(data, col, headers);
    const start = headers ? 1 : 0;
    const want = '' + needle;

    for (let r = start, n = data.length; r < n; r++) {
      if ('' + data[r][colIndex] === want) {
        if (ret === 'row')  return data[r];
        if (ret === 'both') return { index: r, row: data[r] };
        return r; // 'index'
      }
    }
    return undefined;
  }

  /** Return ALL matches (string equality). ret: 'row' | 'index' | 'both' */
  function findRowsByColumn(data, col, needle, headers = false, ret = 'row') {
    if (!Array.isArray(data) || !Array.isArray(data[0])) throw new Error('data must be 2D');

    const colIndex = _resolveColIndexEq_(data, col, headers);
    const start = headers ? 1 : 0;
    const want = '' + needle;

    const out = [];
    for (let r = start, n = data.length; r < n; r++) {
      if ('' + data[r][colIndex] === want) {
        out.push(ret === 'row' ? data[r] : (ret === 'both' ? { index: r, row: data[r] } : r));
      }
    }
    return out;
  }

  /* -------- helpers (minimal & fast) -------- */

  function _resolveColIndexEq_(data, col, headers) {
    if (typeof col === 'number') {
      if (col < 0 || col >= data[0].length) throw new Error('Column index out of bounds: ' + col);
      return col;
    }
    if (typeof col === 'string') {
      if (!headers) throw new Error('Header name requires headers=true');
      const hdr = data[0];
      for (let i = 0, w = hdr.length; i < w; i++) if (hdr[i] === col) return i;
      for (let i = 0, w = hdr.length; i < w; i++) if ('' + hdr[i] === col) return i;
      throw new Error('Header not found: ' + col);
    }
    throw new Error('col must be 0-based index or header string');
  }

  // ---------- Small utils ----------
  function to2D(v) {
    if (Array.isArray(v) && Array.isArray(v[0])) return v;
    if (Array.isArray(v)) return v.map(x => [x]);
    return [[v]];
  }

  function flatten(a) {
    const out = [];
    (Array.isArray(a) ? a : [a]).forEach(r => {
      if (Array.isArray(r)) r.forEach(x => out.push(x));
      else out.push(r);
    });
    return out;
  }

  function toBool(x) {
    if (typeof x === "boolean") return x;
    if (x == null) return false;
    const s = String(x).trim().toLowerCase();
    return !(s === "false" || s === "0" || s === "");
  }

  function isPlainObject(o) {
    return o != null && typeof o === "object" && !Array.isArray(o);
  }

  // -------- String helpers --------
  /** Lowercase first char; replace spaces with underscores. */
  function safePropName(s) {
    if (s == null) return s;
    const t = String(s);
    return (t.charAt(0).toLowerCase() + t.slice(1)).replace(/ /g, "");
  }

  // -------- Value helpers --------
  function coerce(v) {
    if (v == null) return v;
    const t = typeof v;
    if (t === 'boolean' || t === 'number' || v instanceof Date) return v;
    if (t === 'string') {
      const s = v.trim();
      if (!s) return s;
      const sl = s.toLowerCase();
      if (sl === 'true')  return true;
      if (sl === 'false') return false;
      if (/^[\+\-]?\d+(\.\d+)?([eE][\+\-]?\d+)?$/.test(s)) {
        const n = Number(s);
        if (Number.isFinite(n)) return n;
      }
      return s;
    }
    return v;
  }

  // -------- Object helpers --------
  function deepCloneSimple(v) {
    if (v == null || typeof v !== 'object') return v;
    if (v instanceof Date) return new Date(v.getTime());
    if (Array.isArray(v)) return v.map(deepCloneSimple);
    const o = {};
    for (const k of Object.keys(v)) o[k] = deepCloneSimple(v[k]);
    return o;
  }

  function deepEqualSimple(a, b) {
    if (a === b) return true;
    if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
    if (!a || !b || typeof a !== 'object' || typeof b !== 'object') return false;
    if (Array.isArray(a) !== Array.isArray(b)) return false;

    if (Array.isArray(a)) {
      if (a.length !== b.length) return false;
      for (let i = 0; i < a.length; i++) if (!deepEqualSimple(a[i], b[i])) return false;
      return true;
    }
    const ak = Object.keys(a), bk = Object.keys(b);
    if (ak.length !== bk.length) return false;
    for (const k of ak) if (!deepEqualSimple(a[k], b[k])) return false;
    return true;
  }

  function dateToSerial(d) {
    const MS_PER_DAY = 24 * 60 * 60 * 1000;
    return d.getTime() / MS_PER_DAY + 25569; // 1899-12-30 epoch
  }

  function formatDate(d) {
    return Utilities.formatDate(d, EDContext.context.date.tz || Session.getScriptTimeZone() || 'UTC', EDContext.context?.config?.date?.format || "yyyy-MM-dd hh:mm:ss");
  }

  function byteLen(str) { return Utilities.newBlob(str).getBytes().length; }

  /* ===================== NEW: array-of-object search + index ===================== */

  // Internal: fast path getter with '.' or '|' separator + safePropName
  function _getAtPath_(obj, path, sep) {
    if (!path) return obj;
    const parts = String(path).split(sep).map(s => s.trim()).filter(Boolean);
    let cur = obj;
    for (let i = 0; i < parts.length; i++) {
      if (cur == null) return undefined;
      const seg = safePropName(parts[i]);
      cur = cur[seg];
    }
    return cur;
  }

  /**
   * Build an index for an array of objects by a property path.
   * @param {Object[]} arr
   * @param {string} propPath       e.g. 'id', 'user.email', 'user|email'
   * @param {{ret?:'obj'|'index', caseInsensitive?:boolean, stringKeys?:boolean, sep?:'|'|'.'}} [opts]
   * @returns {{getFirst:(needle:any)=>any|number|undefined, getAll:(needle:any)=>Array<any|number>, size:number}}
   */
  function buildIndexBy(arr, propPath, opts = {}) {
    const { ret = 'obj', caseInsensitive = false, stringKeys = true, sep } = opts;
    const useSep = sep || (propPath.indexOf('|') >= 0 ? '|' : '.');

    const keyFn = (v) => {
      if (!stringKeys) return v;
      let s = v == null ? '' : String(v);
      if (caseInsensitive) s = s.toLowerCase();
      return s;
    };
    const valFn = (row, idx) => (ret === 'index' ? idx : row);

    const first = new Map();
    const all = new Map();

    for (let i = 0; i < (arr?.length || 0); i++) {
      const row = arr[i];
      if (row == null) continue;
      const k = keyFn(_getAtPath_(row, propPath, useSep));
      if (!first.has(k)) first.set(k, valFn(row, i));
      const bucket = all.get(k);
      if (bucket) bucket.push(valFn(row, i));
      else all.set(k, [valFn(row, i)]);
    }

    const normKey = (needle) => keyFn(needle);

    return {
      size: arr?.length || 0,
      getFirst: (needle) => first.get(normKey(needle)),
      getAll:   (needle) => all.get(normKey(needle)) || []
    };
  }

  /**
   * Find the first match by property path. Linear scan (use buildIndexBy for repeated lookups).
   * @param {Object[]} arr
   * @param {string} propPath
   * @param {any} needle
   * @param {{ret?:'obj'|'index', caseInsensitive?:boolean, sep?:'|'|'.'}} [opts]
   */
  function findFirst(arr, propPath, needle, opts = {}) {
    const { ret = 'obj', caseInsensitive = false, sep } = opts;
    const useSep = sep || (propPath.indexOf('|') >= 0 ? '|' : '.');

    const want = (v) => v == null ? '' : String(v);
    for (let i = 0; i < (arr?.length || 0); i++) {
      const row = arr[i];
      if (row == null) continue;
      const got = _getAtPath_(row, propPath, useSep);
      if (typeof got === 'string' && typeof needle === 'string' && caseInsensitive) {
        if (got.toLowerCase() === needle.toLowerCase()) return ret === 'index' ? i : row;
      } else if (got === needle || want(got) === want(needle)) {
        return ret === 'index' ? i : row;
      }
    }
    return undefined;
  }

  /**
   * Find ALL matches by property path. Linear scan (use buildIndexBy(...).getAll for repeated).
   * @param {Object[]} arr
   * @param {string} propPath
   * @param {any} needle
   * @param {{ret?:'obj'|'index', caseInsensitive?:boolean, sep?:'|'|'.'}} [opts]
   */
  function findAll(arr, propPath, needle, opts = {}) {
    const { ret = 'obj', caseInsensitive = false, sep } = opts;
    const useSep = sep || (propPath.indexOf('|') >= 0 ? '|' : '.');

    const out = [];
    const want = (v) => v == null ? '' : String(v);

    for (let i = 0; i < (arr?.length || 0); i++) {
      const row = arr[i];
      if (row == null) continue;
      const got = _getAtPath_(row, propPath, useSep);
      if (typeof got === 'string' && typeof needle === 'string' && caseInsensitive) {
        if (got.toLowerCase() === needle.toLowerCase()) out.push(ret === 'index' ? i : row);
      } else if (got === needle || want(got) === want(needle)) {
        out.push(ret === 'index' ? i : row);
      }
    }
    return out;
  }

  function hash32(str) {
    // djb2 (uint32)
    var h = 5381 >>> 0;
    for (var i = 0; i < str.length; i++) {
      h = (((h << 5) + h) + str.charCodeAt(i)) >>> 0;
    }
    return h.toString(16);
  }

  function isLeaf(node) {
    return !!(node && typeof node === 'object' && typeof node.range === 'string');
  }

  return {
    Obj:  { deepCloneSimple, deepEqualSimple, isPlainObject, safePropName, isLeaf },
    Arr:  { to2D, flatten, findRowByColumn, findRowsByColumn, buildIndexBy, findFirst, findAll },
    Str:  { toBool, coerce, byteLen, hash32 },
    Date: { dateToSerial, formatDate }
  };
})();
