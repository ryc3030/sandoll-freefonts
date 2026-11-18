(() => {
  "use strict";
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const state = { fonts: [], web: {}, cov: {}, filter: "all", q: "", sort: "default", sample: "", size: 34, favs: new Set() };
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* 저장 불가 환경 */ } },
  };
  state.favs = new Set(store.get("ff:favs", []).map(String));
  state.sample = store.get("ff:sample", $("#sample").value);
  state.size = store.get("ff:size", 34);

  // 웹폰트 CSS는 카드가 화면에 들어올 때만 불러온다.
  const loaded = new Set();
  function loadCss(href) {
    if (!href || loaded.has(href)) return;
    loaded.add(href);
    const l = document.createElement("link");
    l.rel = "stylesheet"; l.href = href; l.crossOrigin = "anonymous";
    document.head.appendChild(l);
  }
  const io = "IntersectionObserver" in window ? new IntersectionObserver((ents) => {
    for (const e of ents) if (e.isIntersecting) { loadCss(e.target.dataset.css); io.unobserve(e.target); }
  }, { rootMargin: "400px" }) : null;

  const TAGKO = { formal: "격식", neutral: "중립", friendly: "친근", playful: "발랄", cute: "귀여움", handwritten: "손글씨",
    retro: "복고", elegant: "우아", traditional: "전통", academic: "학술", tech: "기술", impact: "강한 제목", calm: "차분" };
  // 상업용: 산돌 '상업용' 태그 또는 라이선스 요약표에서 인쇄·웹·포장·영상이 모두 '사용 가능'
  const commercialOK = (f) => (f.tags || []).includes("상업용") ||
    ((f.licenseSummary || []).length > 0 && f.licenseSummary.filter((r) => /인쇄|웹|포장|영상/.test(r.category)).every((r) => /가능/.test(r.allowed) && !/불가/.test(r.allowed)));
  // 공공 보고서 기호(□ ○ ※ Ⅰ ①)와 한글 11,172자를 모두 지원하는가(glyph-coverage.json 기준)
  const REPORT_SYMS = "□○※ⅠⅡ①";
  const covOf = (f) => state.cov[String(f.id)] || state.cov["G:" + (f.nameEn || "")] || null;
  const reportOK = (f) => { const c = covOf(f); return !!c && c.hangul >= 11172 && ![...REPORT_SYMS].some((ch) => c.missing.includes(ch)); };
  const isCond = (t) => /조건부/.test(t);
  function tagHtml(f) {
    const t = [];
    for (const x of f.tags || []) t.push(`<span class="tag ${x === "OFL" ? "ofl" : ""} ${isCond(x) ? "cond" : ""}">${esc(x.replace("•", "·"))}</span>`);
    if (f.isNew) t.push(`<span class="tag new">신규</span>`);
    if (f.source === "google" || f.alsoGoogle) t.push(`<span class="tag src-g">Google Fonts</span>`);
    if (reportOK(f)) t.push(`<span class="tag rep" title="□ ○ ※ Ⅰ ① 와 한글 11,172자 모두 지원">보고서 기호 OK</span>`);
    return t.join("");
  }
  const meta = (f) => [f.maker, f.category?.replace(/,/g, " · "), f.weights, f.year].filter(Boolean).join(" · ");
  const sampleText = () => state.sample.trim() || "다람쥐 헌 쳇바퀴에 타고파";

  function previewHtml(f, big = false) {
    const w = state.web[f.id];
    if (w) {
      return big
        ? `<textarea aria-label="직접 입력해 보기" style="font-family:'${esc(w.family)}'">${esc(sampleText())}</textarea>`
        : `<div class="live" style="font-family:'${esc(w.family)}', var(--ui);font-size:${state.size}px">${esc(sampleText())}</div><span class="src on">실시간</span>`;
    }
    // 공개 웹폰트가 없는 폰트는 미리보기 대신 안내만 보여 준다(다른 사이트의 견본 이미지는 쓰지 않음)
    const where = f.makerUrl ? "제작사 페이지" : "산돌구름";
    return `<div class="nopv"><p class="nopv-name">${esc(f.name)}</p><p class="nopv-msg">공개 웹폰트가 없어 미리보기를 제공하지 않습니다.<br>모양은 ${where}에서 확인하세요.</p></div>`;
  }

  function match(f) {
    const fl = state.filter;
    if (fl === "ofl" && !f.ofl) return false;
    if (fl === "상업용" && !commercialOK(f)) return false;
    if (["임베딩", "로고•CI"].includes(fl) && !(f.tags || []).includes(fl)) return false;
    if (fl === "report" && !reportOK(f)) return false;
    if (fl === "new" && !(f.sections || []).includes("new")) return false;
    if (fl === "steady" && !(f.sections || []).includes("steady")) return false;
    if (fl === "live" && !state.web[f.id]) return false;
    if (fl === "fav" && !state.favs.has(String(f.id))) return false;
    if (fl === "src-sandoll" && f.source !== "sandoll") return false;
    if (fl === "src-google" && !(f.source === "google" || f.alsoGoogle)) return false;
    if (state.q) {
      const hay = [f.name, f.nameEn, f.slug, f.maker, f.copyright, f.category, f.fileName, ...(f.moods || []), ...(f.tags || [])].join(" ").toLowerCase();
      return state.q.toLowerCase().split(/\s+/).every((w) => hay.includes(w));
    }
    return true;
  }
  function sorted(list) {
    const l = [...list];
    const n = (s) => parseInt(s, 10) || 0;
    if (state.sort === "name") l.sort((a, b) => a.name.localeCompare(b.name, "ko"));
    if (state.sort === "year") l.sort((a, b) => n(b.year) - n(a.year));
    if (state.sort === "weights") l.sort((a, b) => n(b.weights) - n(a.weights));
    return l;
  }

  function render() {
    const list = sorted(state.fonts.filter(match));
    $("#result-count").textContent = `${list.length}종 표시 중`;
    $("#empty").hidden = list.length > 0;
    $("#grid").innerHTML = list.map((f) => {
      const w = state.web[f.id];
      return `<li class="card" data-id="${esc(f.id)}">
        <div class="card-head"><div class="titles"><h2>${esc(f.name)}</h2><p class="meta" title="${esc(meta(f))}">${esc(meta(f))}</p></div>
          <button class="fav" aria-pressed="${state.favs.has(String(f.id))}" aria-label="${esc(f.name)} 찜하기" data-act="fav">♥</button></div>
        <div class="preview" data-act="open" ${w ? `data-css="${esc(w.css)}"` : ""}>${previewHtml(f)}</div>
        <div class="tags">${tagHtml(f)}</div>
        <div class="actions">
          <a class="btn primary" href="${esc(f.makerUrl || f.sandollUrl)}" target="_blank" rel="noopener">${f.source === "google" ? "Google Fonts에서 받기" : f.makerUrl ? "공식 배포처에서 받기" : "산돌구름에서 받기"}</a>
          <button class="btn ghost" data-act="open">라이선스</button>
        </div></li>`;
    }).join("");
    if (io) $$(".preview[data-css]").forEach((el) => io.observe(el)); else $$(".preview[data-css]").forEach((el) => loadCss(el.dataset.css));
  }

  function glyphHtml(f) {
    const c = covOf(f);
    if (!c) return "";
    const miss = c.missing || "";
    const sym = [..."□○※◎ⅠⅡ①②·–→₩㎡℃βα"].map((ch) => `<span class="g ${miss.includes(ch) ? "no" : "ok"}" title="${miss.includes(ch) ? "없음" : "있음"}">${esc(ch)}</span>`).join("");
    return `<h3>기호·한글 지원</h3><p class="glyphs">${sym}</p><p class="note">한글 ${Number(c.hangul).toLocaleString()}자${c.hangul >= 11172 ? "(현대 한글 전부)" : " — 드문 글자는 다른 글꼴로 보일 수 있음"} · 빨간 기호는 이 폰트에 없어 다른 글꼴로 대체됩니다.</p>`;
  }

  function allowedClass(v) { return /불가|금지/.test(v) ? "no" : /조건|문의|허락/.test(v) ? "cond" : /가능/.test(v) ? "ok" : ""; }

  function openDetail(id) {
    const f = state.fonts.find((x) => String(x.id) === String(id));
    if (!f) return;
    const w = state.web[f.id];
    if (w) loadCss(w.css);
    const kv = [["제작사", f.maker], ["저작권", f.copyright], ["분류", f.category], ["종수", f.weights], ["출시", f.year], ["포맷", f.format], ["지원 언어", f.languages], ["글리프", f.spec], ["파일명", f.fileName], ["분위기", (f.moods || []).join(", ")]]
      .filter(([, v]) => v).map(([k, v]) => `<dt>${k}</dt><dd>${esc(v)}</dd>`).join("");
    const sum = (f.licenseSummary || []).map((r) => `<tr><td>${esc(r.category)}</td><td>${esc(r.scope)}</td><td class="${allowedClass(r.allowed)}">${esc(r.allowed)}</td></tr>`).join("");
    const snippet = w ? `<link rel="stylesheet" href="${w.css}">\n\n<style>\n  body { font-family: '${w.family}', sans-serif; }\n</style>` : "";
    $("#d-body").innerHTML = `
      <h2 id="d-name">${esc(f.name)}</h2>
      <p class="sub">${esc([f.nameEn, f.maker].filter(Boolean).join(" · "))}</p>
      <div class="tags" style="padding:12px 0 0">${tagHtml(f)}</div>
      <div class="d-preview">${previewHtml(f, true)}</div>
      ${w ? '<p class="note">위 상자에 직접 글을 써 볼 수 있습니다.</p>' : ""}
      <div class="d-actions">
        ${f.makerUrl ? `<a class="btn primary" href="${esc(f.makerUrl)}" target="_blank" rel="noopener">공식 배포처에서 받기</a>` : ""}
        ${f.sandollUrl ? `<a class="btn ${f.makerUrl ? "ghost" : "primary"}" href="${esc(f.sandollUrl)}" target="_blank" rel="noopener">산돌구름 상세 보기</a>` : ""}
        ${f.googleUrl && f.googleUrl !== f.makerUrl ? `<a class="btn ghost" href="${esc(f.googleUrl)}" target="_blank" rel="noopener">Google Fonts에서 보기</a>` : ""}
      </div>
      <h3>폰트 정보</h3><dl class="kv">${kv}</dl>
      ${glyphHtml(f)}
      ${sum ? `<h3>라이선스 요약 <span class="note" style="font-weight:400">(산돌구름 표기)</span></h3><table class="lic"><thead><tr><th>용도</th><th>범위</th><th>허용</th></tr></thead><tbody>${sum}</tbody></table>` : ""}
      ${f.source === "google" ? `<h3>라이선스</h3><p class="lic-text">${esc(f.license)} — 개인·상업 용도 모두 자유롭게 쓸 수 있고, 수정·재배포도 됩니다. 폰트 파일만 따로 파는 것은 금지됩니다.</p>` : ""}
      ${!f.licenseText && !sum && f.source !== "google" ? `<h3>라이선스</h3><p class="lic-text">산돌구름에 라이선스 정보가 없습니다. 쓰기 전에 제작사에 직접 확인하세요.</p>` : ""}
      ${f.licenseText ? `<h3>라이선스 원문 안내</h3><p class="lic-text">${esc(f.licenseText)}</p>` : ""}
      <p class="note">라이선스는 바뀔 수 있습니다. 실제로 쓰기 전에 배포처의 최신 원문을 확인하세요.</p>
      ${w ? `<h3>웹에서 바로 쓰기</h3><div class="code"><pre>${esc(snippet)}</pre><button class="btn ghost" data-copy>복사</button></div><p class="note">공개 웹폰트 배포처의 CSS를 그대로 연결하는 코드입니다.</p>` : ""}`;
    const cp = $("[data-copy]", $("#d-body"));
    if (cp) cp.addEventListener("click", async () => {
      try { await navigator.clipboard.writeText(snippet); cp.textContent = "복사됨"; } catch { cp.textContent = "복사 실패"; }
      setTimeout(() => (cp.textContent = "복사"), 1500);
    });
    const dlg = $("#detail");
    if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open", "");
    history.replaceState(null, "", `#font-${f.id}`);
  }

  // 이벤트
  $("#grid").addEventListener("click", (e) => {
    const card = e.target.closest(".card"); if (!card) return;
    const id = card.dataset.id;
    const act = e.target.closest("[data-act]")?.dataset.act;
    if (act === "fav") {
      const b = e.target.closest("[data-act=fav]");
      state.favs.has(id) ? state.favs.delete(id) : state.favs.add(id);
      b.setAttribute("aria-pressed", state.favs.has(id));
      store.set("ff:favs", [...state.favs]);
      if (state.filter === "fav") render();
    } else if (act === "open") openDetail(id);
  });
  $$(".chip").forEach((c) => c.addEventListener("click", () => {
    state.filter = c.dataset.filter;
    $$(".chip").forEach((x) => x.setAttribute("aria-pressed", x === c));
    render();
  }));
  let t;
  $("#q").addEventListener("input", (e) => { clearTimeout(t); t = setTimeout(() => { state.q = e.target.value.trim(); render(); }, 120); });
  $("#sort").addEventListener("change", (e) => { state.sort = e.target.value; render(); });
  $("#sample").value = state.sample;
  $("#sample").addEventListener("input", (e) => {
    state.sample = e.target.value; store.set("ff:sample", state.sample);
    $$(".preview .live").forEach((el) => { if (el.nextElementSibling?.classList.contains("on")) el.textContent = sampleText(); });
  });
  const applySize = () => { $("#size-out").textContent = state.size + "px"; $$(".preview .live").forEach((el) => { if (el.nextElementSibling?.classList.contains("on")) el.style.fontSize = state.size + "px"; }); };
  $("#size").value = state.size;
  $("#size").addEventListener("input", (e) => { state.size = Number(e.target.value); store.set("ff:size", state.size); applySize(); });
  $("#detail").addEventListener("close", () => history.replaceState(null, "", location.pathname + location.search));
  $("#detail").addEventListener("click", (e) => { if (e.target === e.currentTarget) e.currentTarget.close(); });

  // 데이터 불러오기
  Promise.all([
    fetch("data/fonts.json").then((r) => r.json()),
    fetch("data/webfonts.json").then((r) => r.json()).catch(() => ({})),
    fetch("data/google-fonts.json").then((r) => r.json()).catch(() => ({ fonts: [] })),
    fetch("data/glyph-coverage.json").then((r) => r.json()).catch(() => ({ fonts: {} })),
  ])
    .then(([data, web, gdata, cov]) => {
      state.cov = cov.fonts || {};
      state.web = Object.fromEntries(Object.entries(web).filter(([k]) => /^\d+$/.test(k)));
      const sandoll = (data.fonts || []).map((f) => ({ ...f, source: "sandoll" }));
      // 산돌구름 폰트 중 Google Fonts와 같은 폰트는 하나로 합친다.
      const byFamily = {};
      for (const f of sandoll) { const w = state.web[f.id]; if (w) byFamily[w.family] = f; }
      const google = [];
      for (const g of gdata.fonts || []) {
        const dup = byFamily[g.family];
        if (dup) { dup.alsoGoogle = true; dup.googleUrl = g.specimenUrl; continue; }
        const id = "g-" + g.family.replace(/\s+/g, "-");
        state.web[id] = { css: g.css, family: g.family };
        google.push({
          id, source: "google", name: g.name || g.family, nameEn: g.family, slug: g.family,
          maker: (g.designers || []).join(", "), category: g.category, weights: (g.weights || []).length + "종",
          year: (g.dateAdded || "").slice(0, 4), tags: g.license === "OFL" ? ["OFL", "상업용", "임베딩", "로고•CI"] : (g.redistributable ? ["상업용", "임베딩"] : []), ofl: g.license === "OFL",
          license: g.license, makerUrl: g.specimenUrl, googleUrl: g.specimenUrl, moods: (g.tags || []).map((t) => TAGKO[t] || t), sections: [],
        });
      }
      state.fonts = [...sandoll, ...google];
      $("#stat-google").textContent = google.length + "종";
      $("#stat-total").textContent = state.fonts.length + "종";
      $("#stat-ofl").textContent = state.fonts.filter((f) => f.ofl).length + "종";
      $("#stat-live").textContent = state.fonts.filter((f) => state.web[f.id]).length + "종";
      if (data.updatedAt) $("#stat-updated").textContent = new Date(data.updatedAt).toLocaleDateString("ko-KR", { year: "numeric", month: "long", day: "numeric" });
      applySize();
      render();
      const m = location.hash.match(/^#font-(.+)/);
      if (m) openDetail(m[1]);
      window.addEventListener("hashchange", () => { const h = location.hash.match(/^#font-(.+)/); if (h) openDetail(h[1]); });
    })
    .catch(() => { $("#result-count").textContent = "목록을 불러오지 못했습니다. 잠시 후 다시 시도해 주세요."; });
})();
