const data = JSON.parse(await Bun.file(import.meta.dir + "/log_timeline.json").text());
const users = [];
let prev = "";
for (const e of data) {
  if (e.role !== "user") continue;
  const t = e.text.trim();
  if (t === prev) continue;
  prev = t;
  if (t.includes("system-reminder")) continue;
  users.push(e);
}
console.log("user messages:", users.length);
for (const e of users) {
  const ts = e.ts.replace("T", " ").slice(0, 16);
  console.log("[" + ts + "] " + t1(e.text));
}
function t1(s) {
  return s.replace(/\s+/g, " ").slice(0, 150);
}
