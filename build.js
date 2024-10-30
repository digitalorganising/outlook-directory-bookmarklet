const fs = require("fs");
const bookmarkleter = require("bookmarkleter");

const code = fs.readFileSync("./get-outlook-365-directory.js", {
  encoding: "utf-8",
});
const bookmarklet = bookmarkleter(code, {
  urlencode: true,
  iife: true,
  minify: true,
  transpile: false,
  jQuery: false,
});
const now = new Date();

const indexTemplate = fs.readFileSync("./index-template.html", {
  encoding: "utf-8",
});
const updatedIndex = indexTemplate
  .replace("REPOSITORY_URL", process.env.REPOSITORY_URL || "#")
  .replace("BOOKMARKLET_TEMPLATE", bookmarklet)
  .replace("LAST_UPDATED_TEMPLATE", now.toISOString())
  .replace(
    "LAST_UPDATED_READABLE_TEMPLATE",
    now.toLocaleString("en-GB", { timeZone: "Europe/London" })
  );

fs.mkdirSync("./build", { recursive: true });
fs.writeFileSync("build/index.html", updatedIndex);
console.log("Wrote updated index.html");
