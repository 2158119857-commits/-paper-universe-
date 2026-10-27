// 用法: bun tools/capture.js <url> <out.jpg> [pointerX,pointerY] [clickSelector] [waitMs]
const EDGE = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
const PORT = 9400 + Math.floor(Math.random() * 400);
const URL_ = process.argv[2];
const OUT = process.argv[3];
const PTR = process.argv[4] || "";
const CLICK = process.argv[5] || "";
const WAIT = +(process.argv[6] || 6000);
const W = +(process.argv[7] || 1280), H = +(process.argv[8] || 800);

const proc = Bun.spawn([EDGE, "--headless=new", "--remote-debugging-port=" + PORT,
  "--user-data-dir=C:\\Users\\21581\\AppData\\Local\\Temp\\edge-cap-" + PORT,
  "--enable-unsafe-swiftshader", "--ignore-gpu-blocklist",
  "--window-size=" + W + "," + H, "--hide-scrollbars", "about:blank"],
  { stdout: "ignore", stderr: "ignore" });

for (let i = 0; i < 80; i++) {
  try { if ((await fetch("http://127.0.0.1:" + PORT + "/json/version")).ok) break; } catch {}
  await new Promise(r => setTimeout(r, 250));
}
const list = await (await fetch("http://127.0.0.1:" + PORT + "/json/list")).json();
const page = list.find(t => t.type === "page");
const ws = new WebSocket(page.webSocketDebuggerUrl);
let id = 0; const pending = new Map();
function send(m, p) { const i = ++id; ws.send(JSON.stringify({ id: i, method: m, params: p || {} })); return new Promise(r => pending.set(i, r)); }
await new Promise(r => ws.addEventListener("open", r));
ws.addEventListener("message", ev => { const m = JSON.parse(ev.data); if (m.id && pending.has(m.id)) { pending.get(m.id)(m.result || m.error); pending.delete(m.id); } });
await send("Page.enable");
await send("Runtime.enable");
await send("Page.navigate", { url: URL_ });
await new Promise(r => setTimeout(r, WAIT));
if (PTR) {
  const [px, py] = PTR.split(",").map(Number);
  await send("Input.dispatchMouseEvent", { type: "mouseMoved", x: px, y: py });
  await new Promise(r => setTimeout(r, 900));
}
if (CLICK) {
  await send("Runtime.evaluate", { expression: "document.getElementById('intro').style.display='none'; (document.querySelector('" + CLICK + "')||{click:function(){}}).click(); 'ok'", returnByValue: true });
  await new Promise(r => setTimeout(r, 4200));
  const [px2, py2] = (PTR || "900,420").split(",").map(Number);
  await send("Input.dispatchMouseEvent", { type: "mouseMoved", x: px2, y: py2 });
  await new Promise(r => setTimeout(r, 1200));
} else if (process.argv[9] !== "keepintro") {
  await send("Runtime.evaluate", { expression: "document.getElementById('intro').style.display='none'" });
  await new Promise(r => setTimeout(r, 600));
}
const shot = await send("Page.captureScreenshot", { format: "jpeg", quality: 86 });
await Bun.write(OUT, Buffer.from(shot.data, "base64"));
console.log("saved", OUT);
try { proc.kill(); } catch {}
process.exit(0);
