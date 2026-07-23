import { copyFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

const pagesDirectory = resolve("docs");
const indexFile = resolve(pagesDirectory, "index.html");

copyFileSync(indexFile, resolve(pagesDirectory, "404.html"));
writeFileSync(resolve(pagesDirectory, ".nojekyll"), "");
