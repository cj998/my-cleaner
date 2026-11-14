/* 정리 엔진 — rules/*.json 의 처방을 읽어 그대로 할 뿐이다.
 * 무엇을 어떻게 고칠지(처방)는 여기에 적지 않는다. 여기에는 «할 줄 아는 일»만 있다.
 *
 *   apply(rows, rules)            처방대로 정리하고, 바꾼 것과 표시한 것을 한 줄씩 남긴다
 *   same(a, b)                    두 결과가 같은가 — «두 번 돌려도 같은 결과» 확인용
 *   compare(rows, answer, key, cols)  정답과 대조
 *
 * 브라우저(window.Clean)와 node(require) 양쪽에서 쓴다. 파일은 밖으로 나가지 않는다.
 */
(function (root) {
  "use strict";

  const FLAG = "_표시";
  const isBlank = v => v === null || v === undefined || String(v).trim() === "";
  const numLike = v => typeof v === "number" || /^\s*-?[\d,]+(\.\d+)?\s*$/.test(String(v));
  const toNum = v => (typeof v === "number" ? v : Number(String(v).replace(/,/g, "").trim()));
  const columnsOf = rows => (rows.length ? Object.keys(rows[0]).filter(c => c !== FLAG) : []);
  const pad2 = n => String(n).padStart(2, "0");
  const daysIn = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
  const isReal = (y, m, d) =>
    Number.isInteger(y) && Number.isInteger(m) && Number.isInteger(d) && m >= 1 && m <= 12 && d >= 1 && d <= daysIn(y, m);
  const show = v => (typeof v === "number" ? v.toLocaleString("ko-KR") : String(v));
  // 엑셀 줄 번호. sheet_to_json 이 남기는 __rowNum__(0부터)이 있으면 그것을 쓴다 — 빈 줄이 끼어도 맞다
  const lineOf = (r, i) => (r && typeof r.__rowNum__ === "number" ? r.__rowNum__ + 1 : i + 2);

  // ── 날짜 — 읽을 줄 아는 표기. 어느 것을 쓸지는 처방의 date_rules[].kind 가 정한다 ──
  const DATE_KINDS = {
    // 2026-05-16 · 2026.05.16 · 2026/05/16
    ymd: {
      re: /^(\d{4})\s*[-.\/]\s*(\d{1,2})\s*[-.\/]\s*(\d{1,2})$/,
      parts: m => ({ y: +m[1], mo: +m[2], d: +m[3] }),
    },
    // 12/18/2026 — 네 자리는 연도. 12보다 큰 쪽이 일, 남은 쪽이 달(소거법)
    slash_mdy: {
      re: /^(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{4})$/,
      parts: (m, dr) => {
        const a = +m[1], b = +m[2], y = +m[3];
        if (a > 12 && b > 12) return { y, mo: NaN, d: NaN };      // 달이 될 수 있는 숫자가 없다
        if (b > 12) return { y, mo: a, d: b };
        if (a > 12) return { y, mo: b, d: a };
        if (dr.both_12_or_less_order === "MD") return { y, mo: a, d: b };
        if (dr.both_12_or_less_order === "DM") return { y, mo: b, d: a };
        return { ambiguous: true };                                // 처방이 정하지 않았으면 고르지 않는다
      },
    },
    // 2026년 6월 19일 · 2026년6월19일
    korean: {
      re: /^(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일$/,
      parts: m => ({ y: +m[1], mo: +m[2], d: +m[3] }),
    },
    // 20260615 — 앞 네 자리는 연도, 뒤 네 자리는 MMDD
    compact: {
      re: /^(\d{4})(\d{2})(\d{2})$/,
      parts: m => ({ y: +m[1], mo: +m[2], d: +m[3] }),
    },
  };
  // 25/7/23 · 7/23/25 처럼 연도가 두 자리라 뜻이 갈리는 표기
  const TWO_DIGIT_YEAR = /^(\d{2}\s*[-.\/]\s*\d{1,2}\s*[-.\/]\s*\d{1,2}|\d{1,2}\s*[-.\/]\s*\d{1,2}\s*[-.\/]\s*\d{2})$/;

  // 반환: { value, flag?, check? } — flag 가 있으면 값은 그대로 두고 표시만 한다
  function readDate(v, r) {
    const s = String(v).trim();
    const rules = r.date_rules || [{ kind: "ymd" }];
    for (const dr of rules) {
      const k = DATE_KINDS[dr.kind];
      if (!k) continue;
      const m = s.match(k.re);
      if (!m) continue;
      const p = k.parts(m, dr);
      if (p.ambiguous) {
        return r.ambiguous === "flag" ? { value: v, flag: r.ambiguous_label || "뜻이 갈리는 날짜" } : { value: v };
      }
      if (!isReal(p.y, p.mo, p.d)) {
        if (dr.impossible === "check") return { value: v, flag: r.impossible_label || "확인 필요: 있을 수 없는 날짜", check: true };
        if (dr.impossible === "flag") return { value: v, flag: r.impossible_label || "있을 수 없는 날짜" };
        return { value: v };
      }
      return { value: `${p.y}-${pad2(p.mo)}-${pad2(p.d)}` };
    }
    if (r.ambiguous === "flag" && TWO_DIGIT_YEAR.test(s)) return { value: v, flag: r.ambiguous_label || "뜻이 갈리는 날짜" };
    return { value: v };
  }

  // ── 칸 사이 조건 ─────────────────────────────────────────────
  function test(row, [col, op, value]) {
    const v = row[col];
    if (op === "blank") return isBlank(v);
    if (isBlank(v)) return false;          // 빈칸은 0이 아니다 — 빈칸을 0으로 읽으면 없는 사실이 생긴다
    const n = typeof value === "number" ? toNum(v) : String(v);
    switch (op) {
      case "==": return n === value;
      case "!=": return n !== value;
      case ">": return toNum(v) > value;
      case "<": return toNum(v) < value;
      case ">=": return toNum(v) >= value;
      case "<=": return toNum(v) <= value;
      default: return false;
    }
  }

  // ── 처방 적용 ────────────────────────────────────────────────
  function apply(input, rules) {
    const log = [];
    const cols = columnsOf(input);
    const idCol = cols[0];
    const lines = input.map(lineOf);
    const ids = input.map(r => r[idCol]);
    const rows = input.map(r => { const o = {}; cols.forEach(c => { o[c] = r[c]; }); return o; });
    const flags = rows.map(() => []);
    const checks = rows.map(() => false);

    // 1) 칸마다 — 처방 파일에 적힌 차례대로
    for (const [col, r] of Object.entries(rules.columns || {})) {
      if (!cols.includes(col)) continue;
      rows.forEach((row, i) => {
        let v = row[col];
        const before = v;
        if (r.trim && typeof v === "string") v = v.trim();
        if (r.map && Object.prototype.hasOwnProperty.call(r.map, String(v))) v = r.map[String(v)];
        if (r.blank_to !== undefined && isBlank(v)) v = r.blank_to;
        if (r.strip_units && typeof v === "string") r.strip_units.forEach(u => { v = v.split(u).join("").trim(); });
        if (r.number && !isBlank(v) && numLike(v)) v = toNum(v);
        if (r.date && !isBlank(v)) {
          const d = readDate(v, r);
          v = d.value;
          if (d.flag) flags[i].push(`${col} ${d.flag}`);
          if (d.check) checks[i] = true;
        }
        if (r.rescale && typeof v === "number" && v > r.rescale.over) v = v / r.rescale.divide;

        // 글자 «12»가 숫자 12가 된 것처럼 겉보기가 같으면 기록하지 않는다 — 기록은 사람이 읽을 것만
        const sameNumber = !isBlank(v) && !isBlank(before) && numLike(v) && numLike(before) && toNum(v) === toNum(before);
        if (String(v) !== String(before) && !sameNumber && !(isBlank(v) && isBlank(before))) {
          log.push({ row: lines[i], id: ids[i], col, from: before, to: v, rule: r.why || "" });
        }
        row[col] = v;

        if (r.missing === "flag" && isBlank(v)) flags[i].push(`${col} 빈칸`);
        if (r.range && r.range.action === "flag" && typeof v === "number") {
          const { min, max } = r.range;
          const out = min !== undefined && v < min ? `${show(min)} 미만` : max !== undefined && v > max ? `${show(max)} 초과` : null;
          if (out) flags[i].push(`${col} ${out}(${show(v)})`);
        }
      });
    }

    // 2) 칸 사이 조건 — 모양을 맞춘 뒤에 본다
    for (const cc of rules.cross_checks || []) {
      rows.forEach((row, i) => {
        if (cc.when.every(w => test(row, w))) { flags[i].push(cc.label); if (cc.check) checks[i] = true; }
      });
    }

    // 3) 완전히 같은 줄 — 모양을 다 맞춘 뒤에 센다
    const keep = rows.map(() => true);
    let dropped = 0;
    if (rules.duplicates && rules.duplicates.action === "drop") {
      const seen = new Map();
      rows.forEach((row, i) => {
        const k = JSON.stringify(cols.map(c => row[c]));
        if (seen.has(k)) {
          keep[i] = false; dropped += 1;
          log.push({ row: lines[i], id: ids[i], col: "(행 전체)", from: "", to: "지움",
            rule: `${rules.duplicates.why || "완전히 같은 행"} — ${seen.get(k)}줄과 같음` });
        } else seen.set(k, lines[i]);
      });
    }

    const kept = [], flagged = [];
    rows.forEach((row, i) => {
      if (!keep[i]) return;
      row[FLAG] = flags[i].join(" · ");
      kept.push(row);
      if (flags[i].length) flagged.push({ row: lines[i], id: ids[i], text: row[FLAG], check: checks[i] });
    });

    const byColumn = {};
    log.forEach(l => { byColumn[l.col] = (byColumn[l.col] || 0) + 1; });
    return {
      rows: kept, log, flagged, idCol, columns: cols,
      summary: {
        before: input.length, after: kept.length, dropped,
        changes: log.length - dropped, flagged: flagged.length,
        checks: flagged.filter(f => f.check).length, byColumn,
      },
    };
  }

  function same(a, b) {
    if (a.length !== b.length) return false;
    return a.every((r, i) => JSON.stringify(r) === JSON.stringify(b[i]));
  }

  /* 정답과 대조 — 정답은 «완성 기준»이 아니라 «내 처방이 무엇을 놓쳤나»를 보는 거울이다.
   * key 로 같은 행을 찾고, cols 의 값을 비교한다. 숫자는 값으로 비교한다. */
  function compare(rows, answer, key, cols) {
    const norm = v => (isBlank(v) ? "" : String(v).trim());
    const k = r => key.map(c => norm(r[c])).join("|");
    const want = new Map(answer.map(r => [k(r), r]));
    const mine = new Set(rows.map(k));
    let match = 0, notInAnswer = 0;
    const diff = {}, examples = [];
    rows.forEach(r => {
      const a = want.get(k(r));
      if (!a) { notInAnswer += 1; return; }
      let ok = true;
      for (const c of cols) {
        const x = r[c], y = a[c];
        const eq = norm(x) === norm(y) || (numLike(x) && numLike(y) && !isBlank(x) && !isBlank(y) && toNum(x) === toNum(y));
        if (!eq) {
          ok = false; diff[c] = (diff[c] || 0) + 1;
          if (examples.length < 12) examples.push({ key: k(r), col: c, got: x, want: y });
        }
      }
      if (ok) match += 1;
    });
    const notInMine = answer.filter(r => !mine.has(k(r))).length;
    return { answerRows: answer.length, myRows: rows.length, match, notInAnswer, notInMine, diff, examples };
  }

  const api = { apply, same, compare, readDate, FLAG, isBlank };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Clean = api;
})(this);
