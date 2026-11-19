(() => {
  "use strict";
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  // 요청 문장 복사
  const copied = document.getElementById("copied");
  document.querySelectorAll(".prompt").forEach((b) => {
    b.addEventListener("click", async () => {
      const text = b.childNodes[b.childNodes.length - 1].textContent.trim();
      try { await navigator.clipboard.writeText(text); copied.textContent = `복사했습니다: "${text}"`; }
      catch { copied.textContent = "복사하지 못했습니다. 문장을 직접 선택해 복사하세요."; }
    });
  });

  // 전후 비교 예시
  const list = document.getElementById("ex-list");
  const IMG = "docs/examples/";
  const regName = (r) => { const n = Array.isArray(r) && r.length ? r[0][0] : ""; return /보류/.test(n) ? "" : n; };
  const why = (e) => {
    if (/내용 우선/.test(e.way)) return `만든 곳은 공공기관이지만 내용이 '${esc(e.genre)}' 성격이 뚜렷해 그에 맞는 서체를 골랐습니다.`;
    if (/공공/.test(e.way)) return "공공기관의 정책 보고 자료라 공공 기본값인 프리텐다드를 골랐습니다. □ ○ ※ Ⅰ 기호를 모두 갖춘 폰트입니다.";
    if (/기업/.test(e.genre + e.way)) return "회사의 실적 보고 자료라 숫자와 표가 잘 읽히는 프리텐다드를 골랐습니다.";
    return "";
  };

  function card(e) {
    const pages = Array.from({ length: e.pages }, (_, i) => i + 1);
    const ag = (e.agency || []).map((a) =>
      `<li>${esc(a.agency)}: <a href="${esc(a.url)}" target="_blank" rel="noopener">${esc(a.fonts.join(", "))}</a> — 파일을 첨부하면 그 파일로 적용</li>`).join("");
    return `<article class="ex" data-key="${esc(e.key)}">
      <header class="ex-head">
        <h3>${esc(e.title)}</h3>
        <p class="sub">만든 곳: ${esc(e.author)}</p>
      </header>
      <div class="ex-body">
        <div class="ex-shots">
          <div class="seg" role="group" aria-label="${esc(e.title)} 보기 전환">
            <button type="button" data-v="before" aria-pressed="false">원본 (맑은 고딕)</button>
            <button type="button" data-v="after" aria-pressed="true">적용 후</button>
          </div>
          ${pages.map((n) => `<img loading="lazy" width="960" height="540" src="${IMG}${esc(e.key)}-after-${n}.jpg" data-n="${n}" alt="${esc(e.title)} ${n}쪽 적용 후">`).join("")}
        </div>
        <div class="ex-diag">
          <p class="ex-label">요청</p>
          <p class="ex-ask">${(e.polish || []).length > 3 ? "어울리는 무료 폰트로 바꾸고, 그래프랑 디자인도 같이 손봐 줘" : "이 피피티에 어울리는 무료 폰트로 바꿔 줘"}</p>
          <p class="ex-label">스킬의 답</p>
          <dl class="kv">
            <dt>진단</dt><dd>${esc(e.genre)}${regName(e.register) ? ` · 문체 ${esc(regName(e.register))}` : ""}</dd>
            <dt>제목</dt><dd><b>${esc(e.titleFont)}</b></dd>
            <dt>본문</dt><dd><b>${esc(e.bodyFont)}</b></dd>
            <dt>대안</dt><dd>${(e.alt || []).map(esc).join("<br>")}</dd>
          </dl>
          <p class="ex-why">${why(e)}</p>
          ${ag ? `<p class="ex-label">기관 전용 서체 제안</p><ul class="ex-ag">${ag}</ul>` : ""}
          ${(e.polish || []).length ? `<p class="ex-label">그래프·디자인 보정</p><ul class="ex-ag">${e.polish.map((m) => `<li>${esc(m)}</li>`).join("")}</ul>` : ""}
        </div>
      </div>
    </article>`;
  }

  fetch(IMG + "examples.json")
    .then((r) => { if (!r.ok) throw new Error(r.status); return r.json(); })
    .then((data) => {
      list.innerHTML = data.map(card).join("");
      list.querySelectorAll(".ex").forEach((ex) => {
        const key = ex.dataset.key;
        ex.querySelectorAll(".seg button").forEach((b) => b.addEventListener("click", () => {
          ex.querySelectorAll(".seg button").forEach((x) => x.setAttribute("aria-pressed", String(x === b)));
          ex.querySelectorAll("img").forEach((img) => {
            img.src = `${IMG}${key}-${b.dataset.v}-${img.dataset.n}.jpg`;
            img.alt = img.alt.replace(/(원본|적용 후)$/, b.dataset.v === "after" ? "적용 후" : "원본");
          });
        }));
      });
    })
    .catch(() => { list.innerHTML = '<p class="note">예시를 불러오지 못했습니다. 잠시 뒤 새로고침해 주세요.</p>'; });
})();
