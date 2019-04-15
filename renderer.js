const xlsx = require("node-xlsx");
const fs = require("fs");

const input = document.querySelector("#fileInput");

input.addEventListener("change", () => {
  const file = input.files[0];
  console.log({ file });
  const { path } = file;

  const spreadsheetData = getSpreadsheetData({ spreadsheetPath: path });

  traverseArray(spreadsheetData);

  const departmentName = spreadsheetData[1][1];
  const assets = getAssets({ spreadsheetData });
  const assetsCount = getAssetsNum({ spreadsheetData });
  const assetsPositionsThatExist = getAssetsPositionThatExist({ assetsCount });
  logAssetsThatExist({ assetsPositionsThatExist, assets });
});

/********* HELPERS METHODS **********************/

var TOTAL_COUNT_ROW_NUM = 0;

function getSpreadsheetData({ spreadsheetPath }) {
  const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(spreadsheetPath));
  const { data } = workSheetsFromBuffer[0];
  return data;
}

function getAssets({ spreadsheetData }) {
  const assets = spreadsheetData[15];
  const filteredAssets = assets.filter(asset => asset !== undefined);
  return filteredAssets;
}

function getAssetsNum({ spreadsheetData }) {
  const assetsTotalCount = spreadsheetData[81];
  return assetsTotalCount.filter(
    assetTotal => assetTotal !== undefined && assetTotal !== "Всего"
  );
}

function getAssetsPositionThatExist({ assetsCount }) {
  const assetsExistPositions = [];
  assetsCount.forEach((asset, index) => {
    if (asset !== 0) {
      assetsExistPositions.push(index);
    }
  });
  return assetsExistPositions;
}

function logAssetsThatExist({ assetsPositionsThatExist, assets }) {
  assetsPositionsThatExist.forEach(position => console.log(assets[position]));
}

function getTotalCountRowNum(rowValue, rowIndex) {
  if (typeof rowValue === "string") {
    if (rowValue.includes("всего")) {
      TOTAL_COUNT_ROW_NUM = rowIndex;
    }
  }
}

function traverseArray(array) {
  array.forEach((row, rowIndex) => {
    if (row.length > 0) {
      row.forEach((rowValue, deepRowIndex) => {
        getTotalCountRowNum(rowValue, rowIndex);
        // console.log(rowValue, "|", rowIndex + 1, deepRowIndex + 1);
      });
    }
  });
}
