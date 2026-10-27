// 从 Claude Code 会话记录里抽出人类可读的对话时间线
const P = "C:\\Users\\21581\\AppData\\Roaming\\CherryStudio\\Data\\Agents\\.claude\\projects\\D----\\df753238-9a4e-45b5-94a0-c555e79e9fd0.jsonl";
const text = await Bun.file(P).text();
const lines = text.split("\n").filter(Boolean);
const out = [];
let n = 0;
for (const line of lines) {
  n++;
  let o;
  try { o = JSON.parse(line); } catch { continue; }
  const ts = o.timestamp || "";
  if (o.type === "user" || o.type === "assistant") {
    const msg = o.message || {};
    let blocks = msg.content;
    if (typeof blocks === "string") blocks = [{ type: "text", text: blocks }];
    if (!Array.isArray(blocks)) continue;
    const texts = [];
    for (const b of blocks) {
      if (b && b.type === "text" && b.text) texts.push(b.text);
    }
    if (!texts.length) continue;
    const t = texts.join("\n").trim();
    if (!t) continue;
    if (t.startsWith("<") && t.includes("system-reminder")) continue;
    out.push({ role: o.type, ts, text: t, line: n });
  } else if (o.type === "queue-operation" && o.operation === "enqueue" && o.content) {
    out.push({ role: "user", ts, text: String(o.content).trim(), line: n });
  }
}
console.log("entries:", out.length);
await Bun.write(import.meta.dir + "/log_timeline.json", JSON.stringify(out, null, 1));
for (const e of out.slice(0, 12)) {
  console.log(e.ts, e.role, JSON.stringify(e.text.slice(0, 90)));
}
