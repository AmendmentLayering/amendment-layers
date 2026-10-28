// Serves web/ over https://localhost:3000 using the Office dev certificate.
// The certificate files are read directly: office-addin-dev-certs'
// getHttpsServerOptions() regenerates them whenever its own install check
// fails, which invalidates a CA that was trusted with certutil.
const https = require("https");
const fs = require("fs");
const os = require("os");
const path = require("path");

const CERTS = path.join(os.homedir(), ".office-addin-dev-certs");
const ROOT = path.join(__dirname, "web");
const PORT = 3000;
const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
};

const options = {
  key: fs.readFileSync(path.join(CERTS, "localhost.key")),
  cert: fs.readFileSync(path.join(CERTS, "localhost.crt")),
  ca: fs.readFileSync(path.join(CERTS, "ca.crt")),
};

https
  .createServer(options, (req, res) => {
    const urlPath = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
    const file = path.join(ROOT, urlPath === "/" ? "index.html" : urlPath);
    if (!file.startsWith(ROOT)) {
      res.writeHead(403).end();
      return;
    }
    fs.readFile(file, (err, data) => {
      if (err) {
        res.writeHead(404).end("Not found");
        return;
      }
      res.writeHead(200, {
        "Content-Type": TYPES[path.extname(file)] || "application/octet-stream",
        "Cache-Control": "no-store",
      });
      res.end(data);
    });
  })
  .listen(PORT, () => console.log(`Serving ${ROOT} at https://localhost:${PORT}`));
