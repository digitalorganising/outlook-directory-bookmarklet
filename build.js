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

const indexTemplate = fs.readFileSync("./index-template.html", {
  encoding: "utf-8",
});
const updatedIndex = indexTemplate.replace("BOOKMARKLET_TEMPLATE", bookmarklet);
fs.writeFileSync(`${process.env.GITHUB_WORKSPACE}/index.html, updatedIndex);
