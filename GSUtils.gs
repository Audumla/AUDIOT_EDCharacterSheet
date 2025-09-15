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
    const want = String(needle);

    for (let r = start, n = data.length; r < n; r++) {
      if (String(data[r][colIndex]) === want) {
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
    const want = String(needle);

    const out = [];
    for (let r = start, n = data.length; r < n; r++) {
      if (String(data[r][colIndex]) === want) {
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
      // exact string match first (no allocations)
      for (let i = 0, w = hdr.length; i < w; i++) if (hdr[i] === col) return i;
      // fallback: stringify cells if header row has non-strings
      for (let i = 0, w = hdr.length; i < w; i++) if (String(hdr[i]) === col) return i;
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
    if (x == null) return false;      // <- explicit null/undefined as false
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
    return (t.charAt(0).toLowerCase() + t.slice(1)).replace(/ /g, "_");
  }

  // -------- Value helpers --------
  /** Coerce strings to boolean/number where appropriate; pass through Date/number/boolean. */
  function coerce(v) {
    if (v == null) return v;
    if (typeof v === 'boolean' || typeof v === 'number' || v instanceof Date) return v;
    if (typeof v === 'string') {
      const s = v.trim();
      if (!s) return s;
      const sl = s.toLowerCase();
      if (sl === 'true')  return true;
      if (sl === 'false') return false;
      if (/^[+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(s)) {
        const n = Number(s);
        if (Number.isFinite(n)) return n;
      }
      return s;
    }
    return v;
  }

  // -------- Object helpers --------
  /** Simple deep clone for POJOs/arrays/Date; leaves functions, Maps, Sets, etc. alone. */
  function deepCloneSimple(v) {
    if (v == null || typeof v !== 'object') return v;
    if (v instanceof Date) return new Date(v.getTime());
    if (Array.isArray(v)) return v.map(deepCloneSimple);
    const o = {};
    for (const k of Object.keys(v)) o[k] = deepCloneSimple(v[k]);
    return o;
  }

  /** Simple deep equality for POJOs/arrays/Date. */
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

  function formatDate(d, tz, format = "yyyy-MM-dd'T'HH:mm:ss") {
    return Utilities.formatDate(d, tz || Session.getScriptTimeZone() || 'UTC', format);
  }


  function byteLen(str)  { return Utilities.newBlob(str).getBytes().length; }

  return {
    Obj:  { deepCloneSimple, deepEqualSimple, isPlainObject, safePropName },
    Arr:  { to2D, flatten, findRowByColumn, findRowsByColumn },
    Str:  { toBool, coerce, byteLen },
    Date: { dateToSerial, formatDate }
  };
})();
