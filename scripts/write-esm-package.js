/* eslint-disable @typescript-eslint/no-require-imports */

const fs = require("fs");
const path = require("path");

const outDir = path.join(process.cwd(), "build", "esm");
fs.mkdirSync(outDir, { recursive: true });

fs.writeFileSync(
	path.join(outDir, "package.json"),
	JSON.stringify({ type: "module" }, null, 2),
	"utf8",
);
