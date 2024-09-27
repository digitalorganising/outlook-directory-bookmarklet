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

const readme = fs.readFileSync("./README.md", {
  encoding: "utf-8",
});
const updatedReadme = readme.replace(
  /(\[bookmarklet-ref\]:)(.*)/,
  "$1" + bookmarklet
);
fs.writeFileSync("./README.md", updatedReadme);
