/*
  Helper para capturar grupos visibles en Facebook.

  Uso:
  1. Abre Facebook y busca grupos.
  2. Abre DevTools > Console.
  3. Pega este script y presiona Enter.
  4. El resultado queda copiado al portapapeles en formato:
     Nombre del grupo | URL | 56 mil miembros

  No hace clicks, no publica y no navega: solo lee datos visibles en pantalla.
*/
(() => {
  const MEMBER_WORDS = /miembros?|members?|seguidores?|followers?/i;
  const MEMBER_WITH_UNIT = /(\d+(?:[.,]\d+)?)\s*(mil|k|m|mill[oó]n(?:es)?|miembros?|members?|seguidores?|followers?)/i;
  const MEMBER_PREFIX = /(?:miembros?|members?|seguidores?|followers?)\D{0,16}(\d+(?:[.,]\d+)?)(?:\s*(mil|k|m|mill[oó]n(?:es)?))?/i;
  const BAD_NAME = /^(grupo|public|private|p[uú]blico|privado|unirte|join|joined|ver grupo|view group|miembros?|members?|seguidores?|followers?|publicaciones|posts)\b/i;

  function clean(text) {
    return String(text || "").replace(/\s+/g, " ").trim();
  }

  function cleanName(text) {
    return clean(text).replace(/^(foto\s+del\s+perfil\s+de|foto\s+de\s+perfil\s+de|imagen\s+del\s+perfil\s+de|profile\s+picture\s+of|profile\s+photo\s+of)\s+/i, "");
  }

  function normalizeUrl(href) {
    try {
      const url = new URL(href, location.origin);
      const parts = url.pathname.split("/").filter(Boolean);
      const groupIndex = parts.findIndex((part) => part.toLowerCase() === "groups");
      if (groupIndex === -1 || !parts[groupIndex + 1]) {
        return "";
      }
      return `https://www.facebook.com/groups/${parts[groupIndex + 1]}`;
    } catch {
      return "";
    }
  }

  function parentText(link) {
    let node = link;
    let best = clean(link.innerText || link.textContent || link.getAttribute("aria-label"));
    for (let depth = 0; depth < 7 && node; depth += 1) {
      const text = clean(node.innerText || node.textContent || "");
      if (text.length > best.length && text.length < 2500) {
        best = text;
      }
      node = node.parentElement;
    }
    return best;
  }

  function extractMembers(text) {
    const chunks = text
      .split(/\n|·|•|\|/g)
      .map(clean)
      .filter(Boolean);
    const prioritized = [
      ...chunks.filter((chunk) => MEMBER_WORDS.test(chunk)),
      ...chunks.filter((chunk) => !MEMBER_WORDS.test(chunk)),
    ];

    for (const chunk of prioritized) {
      const direct = chunk.match(MEMBER_WITH_UNIT);
      if (direct && (MEMBER_WORDS.test(chunk) || /mil|k|m|mill/i.test(direct[2] || ""))) {
        return clean(direct[0]);
      }

      const prefixed = chunk.match(MEMBER_PREFIX);
      if (prefixed) {
        return clean([prefixed[1], prefixed[2] || ""].join(" "));
      }
    }
    return "";
  }

  function extractName(link, text) {
    const linkText = cleanName(link.innerText || link.textContent || link.getAttribute("aria-label"));
    if (linkText && !BAD_NAME.test(linkText) && !MEMBER_WORDS.test(linkText) && linkText.split(" ").length <= 14) {
      return linkText;
    }

    const lines = text
      .split(/\n|·|•/g)
      .map(cleanName)
      .filter(Boolean);
    return (
      lines.find((line) => !BAD_NAME.test(line) && !MEMBER_WORDS.test(line) && line.split(" ").length <= 14) || ""
    );
  }

  const rows = [];
  const seen = new Set();

  for (const link of document.querySelectorAll('a[href*="/groups/"]')) {
    const url = normalizeUrl(link.href);
    if (!url || seen.has(url)) {
      continue;
    }

    const text = parentText(link);
    const name = extractName(link, text);
    if (!name) {
      continue;
    }

    const members = extractMembers(text);
    rows.push(`${name} | ${url} | ${members}`);
    seen.add(url);
  }

  const output = rows.join("\n");
  if (typeof copy === "function") {
    copy(output);
  } else if (navigator.clipboard?.writeText) {
    navigator.clipboard.writeText(output).catch(() => {});
  }

  console.log(`Grupos capturados: ${rows.length}`);
  console.log(output || "No encontre grupos visibles. Desplazate un poco y vuelve a ejecutar.");
})();
