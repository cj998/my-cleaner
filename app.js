/* 정리기 화면.
 * 파일은 이 브라우저 안에서만 읽고 계산한다. 서버로 보내는 코드는 없다.
 * 서버에서 받는 것은 처방 파일(rules/)뿐이고, 동료가 파일을 넣기 전에 전부 미리 받아 둔다
 * — 그래서 어떤 파일을 넣었는지(파일 이름)도 서버는 모른다.
 */
(function () {
  "use strict";

  const $ = sel => document.querySelector(sel);
  const FLAG = Clean.FLAG;
  const RULES = {};          // 이름 → 처방 (rules/index.json 에 적힌 것)
  let rulesLoadNote = "";

  const state = { name: "", base: "", rows: null, columns: [], sheets: [], rules: null, rulesFrom: "", result: null };

  // ── 작은 도구 ───────────────────────────────────────────────
  const el = (tag, attrs = {}, ...kids) => {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === "class") e.className = v;
      else if (k === "text") e.textContent = v;
      else e.setAttribute(k, v);
    }
    kids.flat().forEach(k => { if (k !== null && k !== undefined) e.append(k instanceof Node ? k : String(k)); });
    return e;
  };
  const pad2 = n => String(n).padStart(2, "0");
  const fmt = n => Number(n).toLocaleString("ko-KR");
  const shown = v => (v === "" || v === null || v === undefined ? "(빈칸)" : typeof v === "number" ? fmt(v) : `«${v}»`);
  const baseName = name => name.replace(/\.[^.]+$/, "");
  const show = id => { $(id).hidden = false; };
  const hide = id => { $(id).hidden = true; };

  // ── 파일 읽기 ───────────────────────────────────────────────
  // 엑셀 날짜 칸(서식이 날짜인 숫자)은 화면에 보이는 그대로 YYYY-MM-DD 글자로 읽는다.
  // 보이는 값은 같으므로 바꾼 기록에는 남지 않는다.
  function datesToText(ws) {
    for (const addr of Object.keys(ws)) {
      if (addr[0] === "!") continue;
      const c = ws[addr];
      if (c.t === "n" && c.z && XLSX.SSF.is_date(c.z)) {
        const d = XLSX.SSF.parse_date_code(c.v);
        // 서식(z)을 남기면 SheetJS 가 이 글자 칸을 빈칸으로 읽는다 — w 와 함께 지운다
        if (d) { c.t = "s"; c.v = `${d.y}-${pad2(d.m)}-${pad2(d.d)}`; delete c.w; delete c.z; }
      }
    }
  }

  function readBook(buf, name) {
    const isCsv = /\.csv$/i.test(name);
    const wb = isCsv
      ? XLSX.read(new TextDecoder("utf-8").decode(buf).replace(/^﻿/, ""), { type: "string", raw: true })
      : XLSX.read(buf, { type: "array", cellNF: true });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    if (!isCsv) datesToText(ws);
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "", raw: !isCsv });
    const head = XLSX.utils.sheet_to_json(ws, { header: 1, range: 0 })[0] || [];
    return { rows, sheets: wb.SheetNames, sheetName, columns: head.map(String) };
  }

  const readFile = file => new Promise((ok, fail) => {
    const fr = new FileReader();
    fr.onload = () => ok(fr.result);
    fr.onerror = () => fail(fr.error);
    fr.readAsArrayBuffer(file);
  });
  const readText = file => new Promise((ok, fail) => {
    const fr = new FileReader();
    fr.onload = () => ok(fr.result);
    fr.onerror = () => fail(fr.error);
    fr.readAsText(file, "utf-8");
  });

  // ── 처방 파일 미리 받기 ─────────────────────────────────────
  async function loadRules() {
    try {
      const res = await fetch("rules/index.json", { cache: "no-store" });
      if (!res.ok) throw new Error(res.status);
      const idx = await res.json();
      for (const n of idx.files || []) {
        try {
          const r = await fetch(`rules/${encodeURIComponent(n)}.json`, { cache: "no-store" });
          if (r.ok) RULES[n] = await r.json();
        } catch (e) { /* 이 처방만 건너뛴다 */ }
      }
    } catch (e) {
      rulesLoadNote = "처방 목록(rules/index.json)을 읽지 못했습니다. 파일을 컴퓨터에서 바로 열었다면 처방 파일을 직접 골라 주세요.";
    }
  }

  function findRules(base) {
    if (RULES[base]) return RULES[base];
    return Object.values(RULES).find(r => r.file === base) || null;
  }

  // ── 1. 파일 넣기 ────────────────────────────────────────────
  async function onDataFile(file) {
    if (!file) return;
    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) { alert("엑셀(.xlsx) 또는 .csv 파일을 넣어 주세요."); return; }
    try {
      const book = readBook(await readFile(file), file.name);
      Object.assign(state, {
        name: file.name, base: baseName(file.name), rows: book.rows, columns: book.columns,
        sheets: book.sheets, sheetName: book.sheetName, result: null,
      });
    } catch (e) {
      alert("파일을 읽지 못했습니다: " + e.message);
      return;
    }
    renderFileInfo();
    const r = findRules(state.base);
    setRules(r, r ? `rules/${state.base}.json` : "");
    hide("#result"); hide("#compare");
  }

  function renderFileInfo() {
    $("#file-name").textContent = state.name;
    $("#stat-rows").textContent = fmt(state.rows.length);
    $("#stat-cols").textContent = fmt(state.columns.length);
    $("#col-list").textContent = state.columns.join(" · ");
    $("#sheet-note").textContent = state.sheets.length > 1
      ? `시트가 ${state.sheets.length}개입니다. 첫 시트 «${state.sheetName}»만 읽었습니다.`
      : `시트 «${state.sheetName}»`;
    show("#info");
  }

  // ── 2. 처방 보여 주기 ───────────────────────────────────────
  function setRules(rules, from) {
    state.rules = rules;
    state.rulesFrom = from;
    const box = $("#rules-body");
    box.replaceChildren();
    show("#rules");
    if (!rules) {
      $("#no-rules").hidden = false;
      $("#no-rules-msg").textContent = "이 파일에 맞는 처방이 없습니다";
      $("#no-rules-hint").textContent =
        `«${state.base}.json» 처방 파일을 찾지 못했습니다. 처방 파일(.json)을 골라 올려 주세요.` + (rulesLoadNote ? " " + rulesLoadNote : "");
      $("#run").disabled = true;
      $("#rules-meta").textContent = "";
      return;
    }
    $("#no-rules").hidden = true;
    $("#run").disabled = false;
    $("#rules-meta").textContent =
      `${from}${rules.owner ? " · 정한 사람 " + rules.owner : ""}`;
    if (rules.order_note) box.append(el("p", { class: "order-note", text: rules.order_note }));

    const missingCols = Object.keys(rules.columns || {}).filter(c => !state.columns.includes(c));
    if (missingCols.length) box.append(el("p", { class: "warn", text: `이 파일에 없는 칸이라 건너뜁니다: ${missingCols.join(", ")}` }));

    const list = el("div", { class: "rule-list" });
    for (const [col, r] of Object.entries(rules.columns || {})) list.append(ruleCard(col, describeColumn(r)));
    (rules.cross_checks || []).forEach(cc => list.append(ruleCard("칸 사이", [{
      what: `${cc.when.map(([c, op, v]) => `${c} ${op} ${v}`).join(" 그리고 ")} → «${cc.label}» 표시`, why: cc.why,
    }])));
    if (rules.duplicates && rules.duplicates.action === "drop") {
      list.append(ruleCard("행 전체", [{ what: "모양을 다 맞춘 뒤 완전히 같은 줄은 하나만 남긴다", why: rules.duplicates.why }]));
    }
    box.append(list);
  }

  // 처방 한 칸을 사람이 읽는 문장으로 — 처방 파일에 적힌 것만 옮긴다
  function describeColumn(r) {
    const items = [];
    const main = [];
    if (r.trim) main.push("앞뒤 공백을 지운다");
    if (r.map) main.push("바꾼다: " + Object.entries(r.map).map(([a, b]) => `«${a}» → ${b === "" ? "(빈칸)" : `«${b}»`}`).join(", "));
    if (r.blank_to !== undefined) main.push(`빈칸을 «${r.blank_to}»로`);
    if (r.strip_units) main.push(`${r.strip_units.map(u => `«${u}»`).join(" ")} 글자를 뗀다`);
    if (r.number) main.push("숫자로 바꾼다 (쉼표도 뗀다)");
    if (r.rescale) main.push(`${fmt(r.rescale.over)}을 넘으면 ${r.rescale.divide}로 나눈다`);
    if (r.date) {
      main.push(`날짜를 ${r.date_format || "YYYY-MM-DD"} 하나로`);
      (r.date_rules || []).forEach(dr => {
        let t = `${dr.pattern || dr.kind}: ${dr.how || ""}`;
        if (dr.both_12_or_less) t += ` 둘 다 12 이하면 ${dr.both_12_or_less}`;
        main.push(t);
      });
      if ((r.date_rules || []).some(dr => dr.impossible === "check")) {
        main.push(`있을 수 없는 날짜는 고치지 않고 «${r.impossible_label || "확인 필요"}» — 확인이 필요한 데이터로 분류`);
      }
      if (r.ambiguous === "flag") main.push(`«${r.ambiguous_label || "뜻이 갈리는 날짜"}»는 고치지 않고 표시`);
    }
    if (main.length) items.push({ what: main, why: r.why, note: r.note });
    if (r.missing === "flag") items.push({ what: "빈칸은 채우지 않고 표시", why: r.missing_why });
    if (r.range && r.range.action === "flag") {
      const parts = [];
      if (r.range.min !== undefined) parts.push(`${fmt(r.range.min)} 미만`);
      if (r.range.max !== undefined) parts.push(`${fmt(r.range.max)} 초과`);
      items.push({ what: `${parts.join(" 또는 ")}이면 남기고 표시`, why: r.range_why });
    }
    return items;
  }

  function ruleCard(title, items) {
    const card = el("div", { class: "rule" }, el("h4", { text: title }));
    items.forEach(it => {
      const whats = Array.isArray(it.what) ? it.what : [it.what];
      const ul = el("ul", {}, whats.map(w => el("li", { text: w })));
      card.append(ul);
      if (it.why) card.append(el("p", { class: "why" }, el("span", { text: "이유" }), " ", it.why));
      if (it.note) card.append(el("p", { class: "note", text: it.note }));
    });
    return card;
  }

  async function onRulesFile(file) {
    if (!file) return;
    try {
      const rules = JSON.parse(await readText(file));
      if (!rules.columns) throw new Error("columns 가 없습니다");
      setRules(rules, `${file.name} (내 컴퓨터에서 고름)`);
      hide("#result"); hide("#compare");
    } catch (e) {
      alert("처방 파일을 읽지 못했습니다: " + e.message);
    }
  }

  // ── 3. 규칙대로 정리 ────────────────────────────────────────
  function run() {
    if (!state.rows || !state.rules) return;
    const r1 = Clean.apply(state.rows, state.rules);
    const r2 = Clean.apply(state.rows, state.rules);
    const twiceSame = Clean.same(r1.rows, r2.rows) && JSON.stringify(r1.log) === JSON.stringify(r2.log);
    // 정리본을 한 번 더 정리하면 더 바뀌는 것이 없어야 한다
    const again = Clean.apply(r1.rows.map(r => { const o = { ...r }; delete o[FLAG]; return o; }), state.rules);
    state.result = r1;

    const s = r1.summary;
    $("#r-before").textContent = fmt(s.before);
    $("#r-after").textContent = fmt(s.after);
    $("#r-dropped").textContent = fmt(s.dropped);
    $("#r-changes").textContent = fmt(s.changes);
    $("#r-flagged").textContent = fmt(s.flagged);
    $("#r-checks").textContent = s.checks ? `그중 확인 필요 ${fmt(s.checks)}줄` : "확인 필요 0줄";
    const ok = twiceSame && again.summary.changes === 0 && again.summary.dropped === 0;
    $("#r-twice").textContent = ok ? "같음" : "다름";
    $("#r-twice-card").classList.toggle("bad", !ok);
    $("#r-twice-note").textContent = ok
      ? "두 번 돌려도 같은 결과 · 정리본을 다시 정리해도 바뀌는 값 0"
      : `두 번 돌린 결과 ${twiceSame ? "같음" : "다름"} · 다시 정리하면 바뀌는 값 ${again.summary.changes}, 지우는 줄 ${again.summary.dropped}`;

    const by = Object.entries(s.byColumn).filter(([c]) => c !== "(행 전체)");
    $("#r-bycol").textContent = by.length ? "고친 값: " + by.map(([c, n]) => `${c} ${n}`).join(" · ") : "";

    renderLog(r1);
    renderFlagged(r1);
    show("#result");
    show("#compare");
    $("#cmp-out").replaceChildren();
    $("#result").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  const LIMIT = 200;
  function renderLog(r) {
    const tb = $("#log-body");
    tb.replaceChildren();
    $("#log-id").textContent = r.idCol;
    r.log.slice(0, LIMIT).forEach(l => tb.append(el("tr", { class: l.col === "(행 전체)" ? "drop" : "" },
      el("td", { class: "num", text: l.row }), el("td", { text: l.id }), el("td", { text: l.col }),
      el("td", { text: l.col === "(행 전체)" ? "" : shown(l.from) }), el("td", { text: l.col === "(행 전체)" ? "지움" : shown(l.to) }),
      el("td", { class: "why-cell", text: l.rule }))));
    $("#log-more").textContent = r.log.length > LIMIT ? `처음 ${LIMIT}줄만 보입니다. 전체 ${fmt(r.log.length)}줄은 «바꾼 기록»을 내려받아 보세요.` : `모두 ${fmt(r.log.length)}줄`;
  }

  function renderFlagged(r) {
    const tb = $("#flag-body");
    tb.replaceChildren();
    $("#flag-id").textContent = r.idCol;
    r.flagged.slice(0, LIMIT).forEach(f => tb.append(el("tr", { class: f.check ? "check" : "" },
      el("td", { class: "num", text: f.row }), el("td", { text: f.id }), el("td", { text: f.text }))));
    $("#flag-more").textContent = r.flagged.length > LIMIT ? `처음 ${LIMIT}줄만 보입니다. 전체는 «바꾼 기록»의 «표시한 줄» 시트에 있습니다.` : `모두 ${fmt(r.flagged.length)}줄`;
  }

  // ── 내려받기 — .xlsx 로 내보내 엑셀에서 한글이 깨지지 않는다 ────
  function downloadClean() {
    const r = state.result;
    if (!r) return;
    const ws = XLSX.utils.json_to_sheet(r.rows, { header: [...r.columns, FLAG] });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "정리본");
    XLSX.writeFile(wb, `${state.base}_정리본.xlsx`, { compression: true });
  }

  function downloadLog() {
    const r = state.result;
    if (!r) return;
    const s = r.summary;
    const wb = XLSX.utils.book_new();
    const logAoa = [["줄", r.idCol, "칸", "원래 값", "바뀐 값", "이유"],
      ...r.log.map(l => [l.row, l.id, l.col, l.col === "(행 전체)" ? "" : l.from, l.to, l.rule])];
    const flagAoa = [["줄", r.idCol, "확인 필요", "표시"], ...r.flagged.map(f => [f.row, f.id, f.check ? "예" : "", f.text])];
    const sumAoa = [
      ["항목", "값"], ["원본 파일", state.name], ["처방", state.rulesFrom], ["정한 사람", state.rules.owner || ""],
      ["행 (전)", s.before], ["행 (후)", s.after], ["지운 행", s.dropped], ["고친 값", s.changes],
      ["표시한 행", s.flagged], ["확인 필요 행", s.checks],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(logAoa), "바꾼 기록");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(flagAoa), "표시한 줄");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sumAoa), "요약");
    XLSX.writeFile(wb, `${state.base}_바꾼기록.xlsx`, { compression: true });
  }

  // ── 4. 정답과 대조 — 정답 파일도 내 컴퓨터에서 고르고, 브라우저 안에서만 읽는다 ──
  async function onAnswerFile(file) {
    const out = $("#cmp-out");
    if (!file) return;
    if (!state.result) {
      out.replaceChildren(el("p", { class: "warn", text: "먼저 «규칙대로 정리»를 눌러 주세요." }));
      return;
    }
    out.replaceChildren(el("p", { class: "muted", text: `«${file.name}» 읽는 중…` }));
    out.scrollIntoView({ behavior: "smooth", block: "nearest" });
    try {
      compareWith(readBook(await readFile(file), file.name), file.name);
    } catch (e) {
      out.replaceChildren(el("p", { class: "warn", text: `«${file.name}»과 대조하지 못했습니다: ${e.message}` }));
    }
    $("#compare").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function compareWith(ans, fileName) {
    const file = { name: fileName };
    const r = state.result;
    const cfg = state.rules.answer || {};
    const key = cfg.key || [r.idCol];
    const cols = cfg.compare || r.columns.filter(c => !key.includes(c) && ans.columns.includes(c));
    const missingKey = key.filter(k => !ans.columns.includes(k));
    const out = $("#cmp-out");
    out.replaceChildren();
    if (!ans.rows.length) {
      out.append(el("p", { class: "warn", text: `«${fileName}»에 줄이 없습니다.` }));
      return;
    }
    if (missingKey.length) {
      out.append(el("p", { class: "warn", text: `정답 파일에 열쇠 칸이 없습니다: ${missingKey.join(", ")}` }));
      return;
    }
    const c = Clean.compare(r.rows, ans.rows, key, cols);
    const differ = c.myRows - c.notInAnswer - c.match;
    out.append(
      el("p", { class: "muted", text: `${file.name} · 열쇠: ${key.join(" + ")} · 견준 칸: ${cols.join(", ")}` }),
      el("div", { class: "stats" },
        stat("정답 행", fmt(c.answerRows)),
        stat("일치", fmt(c.match), c.match === c.answerRows && c.myRows === c.answerRows ? "good" : ""),
        stat("어긋난 행", fmt(differ), differ ? "bad" : ""),
        stat("정답에 없는 내 행", fmt(c.notInAnswer), c.notInAnswer ? "bad" : ""),
        stat("내 결과에 없는 정답 행", fmt(c.notInMine), c.notInMine ? "bad" : "")),
    );
    const diffCols = Object.entries(c.diff);
    if (diffCols.length) {
      out.append(el("h4", { text: "어긋난 칸" }),
        el("p", { text: diffCols.map(([col, n]) => `${col} ${fmt(n)}칸`).join(" · ") }));
      const tb = el("tbody");
      c.examples.forEach(x => tb.append(el("tr", {},
        el("td", { text: x.key }), el("td", { text: x.col }), el("td", { text: shown(x.got) }), el("td", { text: shown(x.want) }))));
      out.append(el("div", { class: "table-wrap" }, el("table", {},
        el("thead", {}, el("tr", {}, el("th", { text: key.join("+") }), el("th", { text: "칸" }), el("th", { text: "내 값" }), el("th", { text: "정답" }))), tb)));
    } else {
      out.append(el("p", { class: "good-text", text: "견준 칸에서 어긋난 곳이 없습니다." }));
    }
  }

  function stat(label, value, cls = "") {
    return el("div", { class: `stat ${cls}` }, el("div", { class: "stat-value", text: value }), el("div", { class: "stat-label", text: label }));
  }

  // ── 연결 ────────────────────────────────────────────────────
  // 어디서든 오류가 나면 조용히 멈추지 않고 맨 위에 보여 준다
  function showError(msg) {
    const box = $("#errors");
    box.hidden = false;
    box.append(el("p", { text: "오류: " + msg }));
    box.scrollIntoView({ behavior: "smooth", block: "start" });
  }
  window.addEventListener("error", e => showError(e.message || String(e.error)));
  window.addEventListener("unhandledrejection", e => showError((e.reason && e.reason.message) || String(e.reason)));

  function wire() {
    const drop = $("#drop");
    const pick = $("#pick");
    ["dragenter", "dragover"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("over"); }));
    ["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("over"); }));
    drop.addEventListener("drop", e => onDataFile(e.dataTransfer.files[0]));
    drop.addEventListener("click", () => pick.click());
    drop.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); pick.click(); } });
    pick.addEventListener("change", () => { onDataFile(pick.files[0]); pick.value = ""; });
    // 창 아무 데나 떨어뜨려도 브라우저가 파일을 열어 버리지 않게
    ["dragover", "drop"].forEach(ev => window.addEventListener(ev, e => e.preventDefault()));

    $("#rules-pick").addEventListener("change", e => { onRulesFile(e.target.files[0]); e.target.value = ""; });
    $("#rules-pick2").addEventListener("change", e => { onRulesFile(e.target.files[0]); e.target.value = ""; });
    $("#run").addEventListener("click", run);
    $("#dl-clean").addEventListener("click", downloadClean);
    $("#dl-log").addEventListener("click", downloadLog);
    $("#answer-pick").addEventListener("change", e => { onAnswerFile(e.target.files[0]); e.target.value = ""; });
  }

  wire();
  loadRules().then(() => {
    const n = Object.keys(RULES).length;
    $("#rules-status").textContent = n ? `준비된 처방 ${n}개: ${Object.keys(RULES).join(", ")}` : (rulesLoadNote || "준비된 처방이 없습니다.");
  });
})();
