const ROOT = import.meta.dir;

Bun.serve({
  port: 8080,
  async fetch(req) {
    const url = new URL(req.url);
    let p = decodeURIComponent(url.pathname);
    if (p === "/") p = "/index.html";
    const file = Bun.file(ROOT + p);
    if (await file.exists()) return new Response(file);
    return new Response("404 " + p, { status: 404 });
  },
});

console.log("serving " + ROOT + " at http://localhost:8080");
