// 텔레그램 봇 웹훅 (Cloudflare Worker). 내 채팅(TG_CHAT_ID)에서 온 것만 처리하고, 일은 GitHub Actions에 넘김.
// 시크릿: TG_TOKEN, TG_CHAT_ID, GH_TOKEN(저장소 Actions 쓰기), WEBHOOK_SECRET  /  변수: GH_REPO
// 공지 승인: 버튼 -> 비밀번호 답장 -> 봇이 메시지를 지우고 HMAC(TG_TOKEN, "notice|비밀번호|key|ts")만 넘김 (비밀번호 원문은 어디에도 안 남음)

const PROMPT = "🔐 승인 비밀번호를 이 메시지에 답장으로 보내 주세요";

const HELP = [
  "아래 메뉴 버튼으로 대부분 할 수 있어요. 명령어로도 돼요.",
  "",
  "/status - 최근 자동 작업 결과",
  "/gamagam HAS 2026-11 - HAS 가마감 초안 (YSKD도 같음)",
  "/daily 2026-11 HAS - 가마감용 Daily Report (월, 법인 생략 가능)",
  "/cbr 2026-11 - CBR EUR 일별 환율 채우기 (생략하면 지난달)",
  "/salary 2026-11-19 - 급여환율 수동 실행 (생략하면 오늘)",
  "/meal 2026-12 - 식대 시트 만들기",
  "",
  "공지 발송 대기 알림의 [✅ 승인]을 누르면 비밀번호를 물어봐요. 답장한 비밀번호 메시지는 바로 지워요.",
].join("\n");

export default {
  async fetch(req, env, ctx) {
    if (req.method !== "POST") return new Response("ok");
    if (req.headers.get("X-Telegram-Bot-Api-Secret-Token") !== env.WEBHOOK_SECRET) return new Response("forbidden", { status: 403 });
    const u = await req.json();
    ctx.waitUntil(handle(u, env).catch((e) => say(env, "⚠️ 봇 오류: " + (e && e.message ? e.message : e))));
    return new Response("ok");
  },
};

async function tg(env, method, body) {
  const r = await fetch(`https://api.telegram.org/bot${env.TG_TOKEN}/${method}`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
  });
  return r.json();
}
const say = (env, text, extra = {}) => tg(env, "sendMessage", { chat_id: env.TG_CHAT_ID, text, disable_web_page_preview: true, ...extra });

async function gh(env, path, method = "GET", body) {
  const r = await fetch(`https://api.github.com/repos/${env.GH_REPO}${path}`, {
    method,
    headers: { Authorization: `Bearer ${env.GH_TOKEN}`, Accept: "application/vnd.github+json", "User-Agent": "hasbot", "X-GitHub-Api-Version": "2022-11-28" },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!r.ok) throw new Error(`GitHub ${r.status}: ${(await r.text()).slice(0, 200)}`);
  return r.status === 204 ? null : r.json();
}
const dispatch = (env, wf, inputs = {}) => gh(env, `/actions/workflows/${wf}/dispatches`, "POST", { ref: "main", inputs });

async function hmacHex(key, msg) {
  const k = await crypto.subtle.importKey("raw", new TextEncoder().encode(key), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const sig = await crypto.subtle.sign("HMAC", k, new TextEncoder().encode(msg));
  return [...new Uint8Array(sig)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

function kldNow() { return new Date(Date.now() + 2 * 3600 * 1000); }
function ym(offset) { // 칼리닌그라드 기준 이번 달 + offset 달 -> "YYYY-MM"
  const d = kldNow(); d.setUTCDate(1); d.setUTCMonth(d.getUTCMonth() + offset);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
}
function month(arg) {
  if (!arg) return "";
  let m = arg.match(/^(\d{4})-(\d{1,2})$/);
  if (m) return `${m[1]}-${m[2].padStart(2, "0")}`;
  m = arg.match(/^(\d{1,2})월?$/);
  if (m) return `${kldNow().getUTCFullYear()}-${m[1].padStart(2, "0")}`;
  throw new Error(`월 형식을 모르겠어요: ${arg} (예: 2026-11 또는 11)`);
}
const mm = (s) => (s ? `${Number(s.slice(5))}월` : "");

// ── 작업 실행 (명령어, 메뉴 버튼이 모두 여기로)
async function run(env, job, a = "", b = "") {
  if (job === "gamagam") {
    if (!a || !b) return say(env, "예: /gamagam HAS 2026-11  또는  /gamagam YSKD 11");
    const ent = /^has$/i.test(a) ? "HAS" : "YSKD";
    await dispatch(env, ent === "HAS" ? "gamagam.yml" : "ys_gamagam.yml", { month: b });
    return say(env, `▶ ${ent} ${b} 가마감 초안 실행. 끝나면 파일과 결과를 보내 드려요.`);
  }
  if (job === "daily") {
    const inputs = {};
    if (b) inputs.month = b;
    if (a) inputs.entity = a;
    await dispatch(env, "daily.yml", inputs);
    return say(env, `▶ 가마감용 Daily Report 실행 (${inputs.month || "지난달"}, ${inputs.entity || "HAS, YSKD"})`);
  }
  if (job === "cbr") {
    await dispatch(env, "fx_cbr.yml", b ? { month: b } : {});
    return say(env, `▶ CBR 환율 채우기 실행 (${b || "지난달"})`);
  }
  if (job === "salary") {
    await dispatch(env, "salary_fx.yml", b ? { date: b } : {});
    return say(env, `▶ 급여환율 실행 (기준일 ${b || "오늘"}). 기준일이 아니면 그냥 끝나요.`);
  }
  if (job === "meal") {
    await dispatch(env, "meal.yml", b ? { month: b } : {});
    return say(env, `▶ 식대 시트 실행 (${b || "이번 달"})`);
  }
  if (job === "status") return status(env);
}

async function status(env) {
  const r = await gh(env, "/actions/runs?per_page=30");
  const seen = new Set(), lines = [];
  for (const x of r.workflow_runs) {
    if (/시계|제어판 승인|회사메일/.test(x.name) && x.conclusion === "success") continue;
    if (seen.has(x.name)) continue;
    seen.add(x.name);
    const t = new Date(new Date(x.created_at).getTime() + 2 * 3600 * 1000).toISOString().slice(5, 16).replace("T", " ");
    const icon = x.status !== "completed" ? "⏳" : x.conclusion === "success" ? "✅" : "⚠️";
    lines.push(`${icon} ${t} ${x.name}`);
    if (lines.length >= 10) break;
  }
  return say(env, "최근 작업 (칼리닌그라드 시각)\n" + (lines.join("\n") || "없음"));
}

// ── 메뉴
const MENU = {
  keyboard: [["📊 작업 현황", "📑 가마감"], ["📅 Daily Report", "💱 환율"], ["🍱 식대", "❓ 도움말"]],
  resize_keyboard: true, is_persistent: true, input_field_placeholder: "메뉴를 고르거나 명령어 입력",
};
const btn = (text, data) => ({ text, callback_data: data });
function submenu(name) {
  const p1 = ym(-1), p2 = ym(-2), c = ym(0), n = ym(1);
  if (name === "📑 가마감") return ["가마감 초안 (실행 전에 한 번 더 확인해요)", [
    [btn(`HAS ${mm(p1)}`, `ask:gamagam:HAS:${p1}`), btn(`YSKD ${mm(p1)}`, `ask:gamagam:YSKD:${p1}`)],
    [btn(`HAS ${mm(p2)}`, `ask:gamagam:HAS:${p2}`), btn(`YSKD ${mm(p2)}`, `ask:gamagam:YSKD:${p2}`)]]];
  if (name === "📅 Daily Report") return [`가마감용 Daily Report (${mm(p1)})`, [
    [btn("HAS + YSKD", `run:daily::${p1}`)],
    [btn("HAS만", `run:daily:HAS:${p1}`), btn("YSKD만", `run:daily:YSKD:${p1}`)]]];
  if (name === "💱 환율") return ["환율", [
    [btn(`CBR EUR ${mm(p1)} 채우기`, `run:cbr::${p1}`)],
    [btn("급여환율 오늘 실행", "run:salary::")]]];
  if (name === "🍱 식대") return ["식대 시트", [
    [btn(`${mm(c)} 시트`, `run:meal::${c}`), btn(`${mm(n)} 시트`, `run:meal::${n}`)]]];
  return null;
}
const LABEL = { gamagam: "가마감 초안", daily: "Daily Report", cbr: "CBR 환율", salary: "급여환율", meal: "식대 시트" };

async function handle(u, env) {
  const me = String(env.TG_CHAT_ID);
  if (u.callback_query) {
    const q = u.callback_query;
    if (String(q.from.id) !== me) return;
    const data = q.data || "";
    const [act, ...rest] = data.split(":");
    await tg(env, "answerCallbackQuery", { callback_query_id: q.id });
    const edit = (text, kb) => tg(env, "editMessageText", { chat_id: me, message_id: q.message.message_id, text,
      ...(kb ? { reply_markup: { inline_keyboard: kb } } : {}) });
    if (act === "ok") {
      return say(env, `${PROMPT}\n[${rest[0]}]`, { reply_markup: { force_reply: true, input_field_placeholder: "비밀번호" } });
    }
    if (act === "no") {
      await dispatch(env, "notice_send.yml", { tg_cancel: `${rest[0]}|${Date.now()}` });
      return say(env, `✖ [${rest[0]}] 승인 취소 요청을 보냈어요. 제어판에서 승인한 것이 있어도 발송되지 않아요.`);
    }
    if (act === "ask") {  // 덮어쓰는 작업은 한 번 더 확인
      const [job, a, b] = rest;
      return edit(`${a} ${b} ${LABEL[job]}을(를) 실행할까요? 드라이브의 그달 파일을 덮어써요.`,
        [[btn("▶ 실행", `run:${job}:${a}:${b}`), btn("그만두기", "cancel")]]);
    }
    if (act === "run") {
      const [job, a, b] = rest;
      await edit(`${[LABEL[job], a, b].filter(Boolean).join(" ")} - 실행 요청함`);
      return run(env, job, a, b);
    }
    if (act === "cancel") return edit("그만뒀어요.");
    return;
  }
  const m = u.message;
  if (!m || String(m.chat.id) !== me || String(m.from.id) !== me || !m.text) return;
  const text = m.text.trim();

  // 비밀번호 답장
  const rt = m.reply_to_message;
  if (rt && rt.text && rt.text.startsWith(PROMPT)) {
    await tg(env, "deleteMessage", { chat_id: me, message_id: m.message_id });
    await tg(env, "deleteMessage", { chat_id: me, message_id: rt.message_id });
    const key = (rt.text.match(/\[([^\]]+)\]/) || [])[1];
    if (!key) return say(env, "승인할 항목을 못 찾았어요. 알림의 버튼을 다시 눌러 주세요.");
    const ts = String(Date.now());
    const mac = await hmacHex(env.TG_TOKEN, `notice|${text}|${key}|${ts}`);
    await dispatch(env, "notice_send.yml", { tg_confirm: `${key}|${ts}|${mac}` });
    return say(env, `⏳ [${key}] 승인 요청을 보냈어요 (비밀번호 메시지는 지움). 서버에서 비밀번호를 확인한 뒤 결과를 알려 드려요.`);
  }

  // 메뉴 버튼
  if (text === "📊 작업 현황") return status(env);
  if (text === "❓ 도움말") return say(env, HELP, { reply_markup: MENU });
  const sub = submenu(text);
  if (sub) return say(env, sub[0], { reply_markup: { inline_keyboard: sub[1] } });

  // 명령어
  const [cmd0, ...args] = text.split(/\s+/);
  const cmd = cmd0.replace(/@.*$/, "").toLowerCase();
  const ent = (args.find((x) => /^(has|yskd|ys)$/i.test(x)) || "");
  const entN = ent ? (/^has$/i.test(ent) ? "HAS" : "YSKD") : "";
  const mon = month(args.find((x) => /^\d/.test(x) && !/^\d{4}-\d{2}-\d{2}$/.test(x)) || "");
  const day = args.find((x) => /^\d{4}-\d{2}-\d{2}$/.test(x)) || "";

  if (cmd === "/start" || cmd === "/menu" || text === "메뉴") return say(env, "메뉴를 아래에 띄웠어요.", { reply_markup: MENU });
  if (cmd === "/help") return say(env, HELP, { reply_markup: MENU });
  if (cmd === "/status") return status(env);
  if (cmd === "/gamagam") return run(env, "gamagam", entN, mon);
  if (cmd === "/daily") return run(env, "daily", entN, mon);
  if (cmd === "/cbr") return run(env, "cbr", "", mon);
  if (cmd === "/salary") return run(env, "salary", "", day);
  if (cmd === "/meal") return run(env, "meal", "", mon);
  return say(env, "모르는 명령이에요.\n\n" + HELP, { reply_markup: MENU });
}
