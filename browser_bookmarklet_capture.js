(() => {
  const ENDPOINT = "__CAPTURE_ENDPOINT__";
  const RECEIVER_URL = "__RECEIVER_URL__";
  const TOKEN = "__CAPTURE_TOKEN__";
  const MAX_SCROLLS = __MAX_SCROLLS__;
  const SCROLL_DELAY_MS = __SCROLL_DELAY_MS__;
  const STABLE_ROUNDS = 3;
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

  function wait(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
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
      if (text.length > best.length && text.length < 2200) {
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

  function queryFromPage() {
    try {
      const url = new URL(location.href);
      return clean(url.searchParams.get("q") || "");
    } catch {
      return "";
    }
  }

  function createOverlay() {
    const existing = document.getElementById("fb-group-capture-status");
    if (existing) {
      existing.remove();
    }
    const box = document.createElement("div");
    box.id = "fb-group-capture-status";
    box.style.cssText = [
      "position:fixed",
      "right:18px",
      "bottom:18px",
      "z-index:2147483647",
      "background:#101828",
      "color:white",
      "font:13px/1.35 system-ui,sans-serif",
      "padding:12px 14px",
      "border-radius:8px",
      "box-shadow:0 8px 28px rgba(0,0,0,.28)",
      "max-width:320px",
    ].join(";");
    box.textContent = "Preparando captura...";
    document.body.appendChild(box);
    return box;
  }

  function collectVisible(groupsByUrl) {
    let added = 0;
    for (const link of document.querySelectorAll('a[href*="/groups/"]')) {
      const url = normalizeUrl(link.href);
      if (!url || groupsByUrl.has(url)) {
        continue;
      }

      const text = parentText(link);
      const name = extractName(link, text);
      if (!name) {
        continue;
      }

      groupsByUrl.set(url, {
        n: name,
        u: url,
        m: extractMembers(text),
        r: text.slice(0, 900),
      });
      added += 1;
    }
    return added;
  }

  function postToReceiver(popup, type, data) {
    if (!popup || popup.closed) {
      return;
    }
    const origin = new URL(RECEIVER_URL).origin;
    popup.postMessage({ token: TOKEN, type, ...data }, origin);
  }

  async function sendDirect(payload) {
    const response = await fetch(`${ENDPOINT}?token=${encodeURIComponent(TOKEN)}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    return response.json();
  }

  async function main() {
    const overlay = createOverlay();
    let receiverReady = false;
    const receiverOrigin = new URL(RECEIVER_URL).origin;

    function handleReceiverMessage(event) {
      if (event.origin === receiverOrigin && event.data?.token === TOKEN && event.data?.type === "ready") {
        receiverReady = true;
      }
    }
    window.addEventListener("message", handleReceiverMessage);
    const popup = window.open(RECEIVER_URL, "fb_group_capture", "width=620,height=620");

    const groupsByUrl = new Map();
    let stableRounds = 0;
    for (let step = 0; step <= MAX_SCROLLS; step += 1) {
      const added = collectVisible(groupsByUrl);
      stableRounds = added ? 0 : stableRounds + 1;

      overlay.textContent = `Capturando grupos visibles... ${groupsByUrl.size} encontrados (${step}/${MAX_SCROLLS})`;
      postToReceiver(popup, "progress", { count: groupsByUrl.size, step, maxScrolls: MAX_SCROLLS });

      if (step >= MAX_SCROLLS || stableRounds >= STABLE_ROUNDS) {
        break;
      }

      window.scrollBy({ top: Math.max(500, Math.floor(window.innerHeight * 0.9)), behavior: "smooth" });
      await wait(SCROLL_DELAY_MS);
    }

    const groups = Array.from(groupsByUrl.values());
    if (!groups.length) {
      overlay.textContent = "No encontre grupos visibles. Desplazate un poco y vuelve a intentar.";
      postToReceiver(popup, "error", { message: "No encontre grupos visibles." });
      window.removeEventListener("message", handleReceiverMessage);
      return;
    }

    const payload = {
      q: queryFromPage(),
      title: document.title,
      source: location.href,
      capturedAt: new Date().toISOString(),
      groups,
    };

    overlay.textContent = `Enviando ${groups.length} grupos al Excel...`;
    try {
      for (let attempt = 0; attempt < 12 && popup && !popup.closed && !receiverReady; attempt += 1) {
        await wait(250);
      }

      if (popup && !popup.closed && receiverReady) {
        postToReceiver(popup, "capture", { payload });
        overlay.textContent = `Enviados ${groups.length} grupos. Revisa la ventana local.`;
      } else {
        const result = await sendDirect(payload);
        overlay.textContent = `Captura lista: ${result.counters?.added || 0} nuevos, ${result.counters?.updated || 0} actualizados.`;
      }
    } catch (error) {
      const lines = groups.map((group) => `${group.n} | ${group.u} | ${group.m}`).join("\n");
      navigator.clipboard?.writeText(lines).catch(() => {});
      overlay.textContent = "No pude enviar al servidor local. Deje una copia de respaldo en el portapapeles.";
      console.error(error);
    } finally {
      window.removeEventListener("message", handleReceiverMessage);
    }
  }

  main().catch((error) => {
    console.error(error);
    alert(`Error capturando grupos: ${error.message || error}`);
  });
})();
