function toBool(x) {
  if (typeof x === "boolean") return x;
  if (x == null) return false;
  const s = String(x).trim().toLowerCase();
  return !(s === "false" || s === "0" || s === "");
}

var GSPerf = (function () {
  // WeakMap ensures each original object is wrapped once

  const ENABLED_ID = "PERFORMANCE_ENABLED";
  const enabled = toBool(PropertiesService.getScriptProperties().getProperty(ENABLED_ID) ?? false);

  const WRAP_CACHE = new WeakMap();

  function timeAspect(target, opts = {}, _seen = WRAP_CACHE) {
    const {
      label = '',
      thresholdMs = 0,
      include = null,
      exclude = [],
      deep = false,
    } = opts;

    if (!target || (typeof target !== 'object' && typeof target !== 'function')) return target;
    if (_seen.has(target)) return _seen.get(target);

    const isFn = (v) => typeof v === 'function';
    const isWrappableObject = (v) => v && (typeof v === 'object' || typeof v === 'function')
      && !(v instanceof Date) && !(v instanceof RegExp) && !(v instanceof Promise);

    const shouldWrapMethod = (prop, value) => {
      if (!isFn(value)) return false;
      if (include && !include.includes(prop)) return false;
      if (exclude && exclude.includes(prop)) return false;
      return prop !== 'constructor' && prop !== 'prototype';
    };

    const now = () => (typeof performance !== 'undefined' && performance && typeof performance.now === 'function')
      ? performance.now()
      : Date.now();

    const formatLine = (name, ms, args, err) => {
      const pfx = label ? `[${label}] ` : '';
      const dur = `${ms.toFixed(2)} ms`;
      const argPreview = args.map(previewOne).join(', ');
      return err
        ? `${pfx}${name}(${argPreview}) threw after ${dur}: ${String(err)}`
        : `${pfx}${name}(${argPreview}) ${dur}`;
    };

    const previewOne = (a) => {
      try {
        if (a == null) return String(a);
        if (typeof a === 'string') return JSON.stringify(a.length > 60 ? a.slice(0, 57) + '…' : a);
        if (typeof a === 'number' || typeof a === 'boolean') return String(a);
        if (Array.isArray(a)) return `Array(${a.length})`;
        if (typeof a === 'object') {
          const keys = Object.keys(a);
          return `Object(${keys.slice(0,3).join(',')}${keys.length>3?',…':''})`;
        }
        return typeof a;
      } catch { return '…'; }
    };

    const wrapFn = (fn, name) => function timed(...args) {
      const t0 = now();
      try {
        const ret = fn.apply(this, args);
        if (ret && typeof ret.then === 'function') {
          return ret.then((val) => {
            const ms = now() - t0;
            if (ms >= thresholdMs) EDLogger.perf(formatLine(name, ms, args));
            return deep && isWrappableObject(val) ? timeAspect(val, opts, _seen) : val;
          }).catch((err) => {
            const ms = now() - t0;
            EDLogger.perf(formatLine(name, ms, args, err));
            throw err;
          });
        } else {
          const ms = now() - t0;
          if (ms >= thresholdMs) EDLogger.perf(formatLine(name, ms, args));
          return deep && isWrappableObject(ret) ? timeAspect(ret, opts, _seen) : ret;
        }
      } catch (err) {
        const ms = now() - t0;
        EDLogger.perf(formatLine(name, ms, args, err));
        throw err;
      }
    };

    const handler = {
      get(obj, prop, receiver) {
        const value = Reflect.get(obj, prop, receiver);

        // Wrap methods (cached per function)
        if (shouldWrapMethod(String(prop), value)) {
          const cacheKey = '__timed_wrapped__';
          if (!value[cacheKey]) {
            try {
              Object.defineProperty(value, cacheKey, {
                value: wrapFn(value, String(prop)),
                enumerable: false
              });
            } catch {
              return wrapFn(value, String(prop));
            }
          }
          return value[cacheKey];
        }

        // Deep-wrap object-valued properties
        if (deep && isWrappableObject(value)) {
          return timeAspect(value, opts, _seen);
        }

        return value;
      },

      set(obj, prop, value, receiver) {
        // If assigning an object and deep=true, wrap it before storing
        const toSet = (deep && isWrappableObject(value)) ? timeAspect(value, opts, _seen) : value;
        return Reflect.set(obj, prop, toSet, receiver);
      },

      // Preserve property descriptors where possible
      getOwnPropertyDescriptor(obj, prop) {
        return Reflect.getOwnPropertyDescriptor(obj, prop);
      },

      ownKeys(obj) {
        return Reflect.ownKeys(obj);
      }
    };

    const proxy = new Proxy(target, handler);
    _seen.set(target, proxy);
    return proxy;
  }

  function start() {
    if (EDContext.context.logger?.settings.monitorPerformance ?? enabled) {
      GSBatch = timeAspect(GSBatch);
      EDProperties.path = timeAspect(EDProperties.path);
      EDConfig = timeAspect(EDConfig);
      EDEvent = timeAspect(EDEvent);
    }
  }

  function monitor(target, opts = {}) {

    if (EDContext.context.logger?.settings.monitorPerformance ?? enabled) {
      if (opts?.init ?? false) {
        start();
      }
      return timeAspect(target, opts);
    } else {
      return target;
    }
  }

  function stop() {
    PropertiesService.getScriptProperties().setProperty(ENABLED_ID,EDContext.context.logger?.settings.monitorPerformance ?? enabled);
  }

  return {
    start,
    stop,
    monitor
  };
})();
