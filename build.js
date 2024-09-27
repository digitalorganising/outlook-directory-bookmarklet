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

const index = fs.readFileSync("./index.html", {
  encoding: "utf-8",
});
const updatedIndex = index.replace(
  /(a href=")(.*)(">)/,
  "$1" + bookmarklet + "$3"
);
fs.writeFileSync("./index.html", updatedIndex);
