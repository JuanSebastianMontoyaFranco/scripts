var express = require('express');
var fs = require('fs');
var path = require('path');
var app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const port = 6000;
const hooksDir = path.join(__dirname, 'files');
fs.mkdirSync(hooksDir, { recursive: true });
let hookCounter = 0;

function buildHookFilename() {
   const now = new Date();
   const ts = now.toISOString().replace(/[:.]/g, '-');
   hookCounter = (hookCounter + 1) % 100000;
   return `hook-${ts}-${process.pid}-${hookCounter}.json`;
}

app.all(/.*/, function (req, res) {
   console.log("-------------- New Request --------------");
   const shopDomain = req.get('X-Shopify-Shop-Domain')
   //console.log("Headers:"+ JSON.stringify(req.headers, null, 3));
   console.log("Entrada:"+ JSON.stringify(req.body, null, 3));
   console.log("Dominio:", shopDomain);

   const hookPayload = req.body;
   const hookFile = path.join(hooksDir, buildHookFilename());
   try {
      fs.writeFileSync(hookFile, JSON.stringify(hookPayload, null, 2), 'utf8');
   } catch (err) {
      console.error('No se pudo guardar el hook:', err);
   }
   
   res.json({ message: "Thank you for the message" });
})

const server = app.listen(port, function () {
   console.log(`Servidor iniciado en el puerto ${port}`)
})

// keep the HTTP server handle referenced so the process stays alive
server.ref();

server.on('error', function (err) {
   console.error('Servidor falló:', err);
})
