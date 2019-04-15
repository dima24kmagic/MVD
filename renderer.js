const xlsx = require("node-xlsx");
const fs = require("fs");

const input = document.querySelector("#fileInput");

input.addEventListener("change", e => {
  const path = input.files[0].path;
  console.log({ path });
  const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(path));
  console.log({ workSheetsFromBuffer });
});
