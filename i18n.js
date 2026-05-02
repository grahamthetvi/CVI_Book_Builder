/**
 * Lightweight i18n: fetch locale JSON, interpolate {key}, apply [data-i18n] to the DOM.
 */

export const LOCALE_STORAGE_KEY = "cviBookBuilderLocale";
export const SUPPORTED_LOCALES = [
  { code: "en", label: "English" },
  { code: "ar", label: "العربية" }
];

let bundle = null;
let locale = "en";
const listeners = new Set();

function get(obj, path) {
  return path.split(".").reduce((acc, k) => (acc && acc[k] !== undefined ? acc[k] : undefined), obj);
}

export function getLocale() {
  return locale;
}

export function isRtl() {
  return locale === "ar";
}

export function t(path, params = {}) {
  let s = get(bundle, path);
  if (s === undefined || s === null) {
    console.warn(`[i18n] missing: ${path}`);
    return path;
  }
  if (typeof s !== "string") return String(s);
  return s.replace(/\{(\w+)\}/g, (_, key) =>
    params[key] !== undefined && params[key] !== null ? String(params[key]) : `{${key}}`
  );
}

export function onLocaleChange(fn) {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

function notifyLocaleChange() {
  listeners.forEach((fn) => {
    try {
      fn(locale);
    } catch (e) {
      console.warn(e);
    }
  });
}

function guessInitialLocale() {
  try {
    const saved = localStorage.getItem(LOCALE_STORAGE_KEY);
    if (saved && SUPPORTED_LOCALES.some((x) => x.code === saved)) return saved;
  } catch (e) {
    /* ignore */
  }
  const nav = (navigator.language || "en").slice(0, 2).toLowerCase();
  if (SUPPORTED_LOCALES.some((x) => x.code === nav)) return nav;
  return "en";
}

export async function loadLocale(code) {
  const res = await fetch(`locales/${code}.json`, { cache: "no-store" });
  if (!res.ok) throw new Error(`Locale ${code} failed (${res.status})`);
  bundle = await res.json();
  locale = code;
  try {
    localStorage.setItem(LOCALE_STORAGE_KEY, code);
  } catch (e) {
    /* ignore */
  }
  document.documentElement.lang = code;
  document.documentElement.dir = isRtl() ? "rtl" : "ltr";
  notifyLocaleChange();
}

export async function initI18n() {
  const initial = guessInitialLocale();
  await loadLocale(initial);
}

export async function setLocale(code) {
  if (!SUPPORTED_LOCALES.some((x) => x.code === code)) return;
  await loadLocale(code);
}

/**
 * Apply translations to elements with data-i18n (text), data-i18n-html, data-i18n-placeholder, data-i18n-aria-label, data-i18n-title.
 * Optional data-i18n-params as JSON string for interpolation.
 */
export function applyDomTranslations(root = document) {
  if (!bundle) return;

  root.querySelectorAll("[data-i18n]").forEach((el) => {
    const key = el.getAttribute("data-i18n");
    let params = {};
    const raw = el.getAttribute("data-i18n-params");
    if (raw) {
      try {
        params = JSON.parse(raw);
      } catch (e) {
        /* ignore */
      }
    }
    el.textContent = t(key, params);
  });

  root.querySelectorAll("[data-i18n-html]").forEach((el) => {
    const key = el.getAttribute("data-i18n-html");
    el.innerHTML = t(key);
  });

  root.querySelectorAll("[data-i18n-placeholder]").forEach((el) => {
    const key = el.getAttribute("data-i18n-placeholder");
    el.setAttribute("placeholder", t(key));
  });

  root.querySelectorAll("[data-i18n-aria-label]").forEach((el) => {
    const key = el.getAttribute("data-i18n-aria-label");
    el.setAttribute("aria-label", t(key));
  });

  root.querySelectorAll("[data-i18n-alt]").forEach((el) => {
    const key = el.getAttribute("data-i18n-alt");
    el.setAttribute("alt", t(key));
  });

  root.querySelectorAll("[data-i18n-title]").forEach((el) => {
    const key = el.getAttribute("data-i18n-title");
    el.setAttribute("title", t(key));
  });

  root.querySelectorAll("select[data-i18n-options]").forEach((select) => {
    const prefix = select.getAttribute("data-i18n-options");
    Array.from(select.options).forEach((opt) => {
      const k = opt.getAttribute("data-i18n-opt");
      if (k) opt.textContent = t(`${prefix}.${k}`);
    });
  });
}

export function formatLocaleList(items, style = "narrow", type = "list") {
  try {
    return new Intl.ListFormat(locale, { style, type }).format(items);
  } catch (e) {
    return items.join(", ");
  }
}
