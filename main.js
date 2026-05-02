import {
  initI18n,
  setLocale,
  applyDomTranslations,
  getLocale,
  SUPPORTED_LOCALES,
  onLocaleChange,
  t
} from "./i18n.js";

function refreshStaticTranslations() {
  applyDomTranslations(document);
  const tpl = document.getElementById("spreadTemplate");
  if (tpl && tpl.content) applyDomTranslations(tpl.content);
  document.title = t("document.pageTitle");
}

await initI18n();
refreshStaticTranslations();

const { bootstrap, applyLocaleRefresh } = await import("./app.js");

const localeSelect = document.getElementById("localeSelect");
if (localeSelect) {
  localeSelect.innerHTML = SUPPORTED_LOCALES.map((l) => `<option value="${l.code}">${l.label}</option>`).join("");
  localeSelect.value = getLocale();
  localeSelect.addEventListener("change", async () => {
    await setLocale(localeSelect.value);
    refreshStaticTranslations();
    applyLocaleRefresh();
  });
}

onLocaleChange(() => {
  if (!localeSelect) return;
  localeSelect.value = getLocale();
});

bootstrap();
