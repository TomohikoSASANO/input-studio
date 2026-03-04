;(function setupI18n() {
  const STORAGE_KEY = "inputstudio-locale"
  const SUPPORTED = ["ja", "en", "zh", "hi", "es", "fr", "ar", "pt", "ru", "bn", "id", "ur", "de", "it", "tr", "vi", "ko", "fa", "th", "pl", "uk", "nl"]
  const FALLBACK = "ja"

  const dictionaries = {
    ja: {},
    en: {},
  }

  const interpolate = (text, vars) => {
    return String(text || "").replace(/\{(\w+)\}/g, (_, k) => String(vars?.[k] ?? ""))
  }

  const normalize = (locale) => {
    const v = String(locale || "").toLowerCase()
    if (!v) return FALLBACK
    if (SUPPORTED.includes(v)) return v
    const short = v.split("-")[0]
    return SUPPORTED.includes(short) ? short : FALLBACK
  }

  const getSavedLocale = () => {
    try {
      return normalize(localStorage.getItem(STORAGE_KEY))
    } catch {
      return FALLBACK
    }
  }

  let currentLocale = getSavedLocale()

  const setLocale = (locale) => {
    currentLocale = normalize(locale)
    try {
      localStorage.setItem(STORAGE_KEY, currentLocale)
    } catch {}
    try {
      document.documentElement.lang = currentLocale
    } catch {}
    return currentLocale
  }

  const getLocale = () => currentLocale

  const t = (key, vars = {}) => {
    const k = String(key || "")
    const dic = dictionaries[currentLocale] || {}
    const en = dictionaries.en || {}
    const fb = dictionaries[FALLBACK] || {}
    const raw = dic[k] ?? en[k] ?? fb[k] ?? k
    return interpolate(raw, vars)
  }

  const loadLocaleFile = async (locale) => {
    try {
      const res = await fetch(`./locales/${locale}.json?v=1`, { cache: "no-store" })
      if (!res.ok) return
      const data = await res.json()
      if (data && typeof data === "object") {
        dictionaries[locale] = { ...dictionaries[locale], ...data }
      }
    } catch {}
  }

  const ready = Promise.all(SUPPORTED.map((locale) => loadLocaleFile(locale))).then(() => setLocale(currentLocale))

  window.i18n = {
    t,
    setLocale,
    getLocale,
    ready,
    supported: [...SUPPORTED],
  }

  setLocale(currentLocale)
})()
