var GSJson = (function () {
  /**
   * arrayToJson(input, headers=false)
   * - If Range/A1: uses GSRange.getDisplayArray (display strings, trimmed rows).
   * - headers = TRUE: first row = headers → array of objects.
   * - headers = FALSE (default): array-of-arrays.
   *
   * @param {GoogleAppsScript.Spreadsheet.Range|string|any[][]} input
   * @param {boolean} [headers=false]
   * @return {string} JSON string
   */
  function arrayToJson(input, headers = false) {
    const arr = (typeof input === 'string' || GSUtils.Types.isRangeLike(input))
      ? GSRange.getDisplayArray(input)
      : GSUtils.Arr.to2D(input);

    const useHeaders = GSUtils.Val.toBool(headers);
    if (!useHeaders) return JSON.stringify(arr);
    if (!arr.length) return "[]";

    const hdr = arr[0].map(String);
    const body = arr.length > 1 ? arr.slice(1) : [];
    const objs = body.map(row => {
      const o = {};
      for (let i = 0; i < hdr.length; i++) o[hdr[i]] = row[i];
      return o;
    });
    return JSON.stringify(objs);
  }

  /**
   * jsonToArray(json, keys)
   * - Array of objects → [header; rows...]
   * - Array of arrays  → as-is
   * - 1D array         → single column
   * - Scalar           → 1x1
   *
   * @param {string|any} jsonText
   * @param {string|string[]|any[]} [keys]  // comma list or array (can be nested)
   * @return {any[][]}
   */
  function jsonToArray(jsonText, keys) {
    if (jsonText == null || jsonText === "") return [[""]];

    let data;
    try { data = (typeof jsonText === "string") ? JSON.parse(jsonText) : jsonText; }
    catch (err) { return [["JSON parse error: " + err]]; }

    // Array of objects -> table
    if (Array.isArray(data) && data.length && GSUtils.Val.isPlainObject(data[0])) {
      let order;
      if (keys !== undefined && keys !== "") {
        order = Array.isArray(keys)
          ? GSUtils.Arr.flatten(keys)
          : String(keys).split(",").map(s => s.trim()).filter(Boolean);
      } else {
        const seen = new Set();
        for (let i = 0; i < data.length; i++) {
          const o = data[i];
          if (o && typeof o === "object") for (const k in o) if (!seen.has(k)) seen.add(k);
        }
        order = Array.from(seen);
      }
      const rows = data.map(o => order.map(k => (o && typeof o === "object") ? o[k] : null));
      return [order].concat(rows);
    }

    // Array of arrays
    if (Array.isArray(data) && data.length && Array.isArray(data[0])) return data;

    // 1D array -> column
    if (Array.isArray(data)) return data.map(v => [v]);

    // Scalar
    return [[data]];
  }


  return {
    arrayToJson,
    jsonToArray,
  };
})();
