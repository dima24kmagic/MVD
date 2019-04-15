const xlsx = require("node-xlsx");
const fs = require("fs");

const input = document.querySelector("#fileInput");

input.addEventListener("change", () => {
  const file = input.files[0];
  const { path, name } = file;
  setChosenFiles(name)
  const spreadsheetData = getSpreadsheetData({ spreadsheetPath: path });

  dropArea.style.border = '2px dashed #00FF93'
  dropArea.style.background = 'rgba(0, 255, 147, 0.44)'

  label.style.background = 'rgba(0, 255, 147, 0.44)'

  traverseArrayAndFindRows(spreadsheetData);
  const departmentName = spreadsheetData[1][1];
  const assets = getAssets({ spreadsheetData });
  const assetsCount = getAssetsNum({ spreadsheetData });
  const assetsPositionsThatExist = getAssetsPositionThatExist({ assetsCount });
  logAssetsThatExist({ assetsPositionsThatExist, assets });
});

/********* HELPERS METHODS **********************/

var TOTAL_COUNT_ROW_NUM = 0;
var ASSETS_ROW_NUM = 0;

function getSpreadsheetData({ spreadsheetPath }) {
  const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(spreadsheetPath));
  const { data } = workSheetsFromBuffer[0];
  return data;
}

function getAssets({ spreadsheetData }) {
  const assets = spreadsheetData[ASSETS_ROW_NUM + 1];
  const filteredAssets = assets.filter(asset => asset !== undefined);
  return filteredAssets;
}

function getAssetsNum({ spreadsheetData }) {
  const assetsTotalCount = spreadsheetData[TOTAL_COUNT_ROW_NUM];
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
    if (rowValue.includes("Всего")) {
      TOTAL_COUNT_ROW_NUM = rowIndex;
    }
  }
}
function getAssetsRowNum(rowValue, rowIndex) {
  if (typeof rowValue === "string") {
    if (rowValue.includes("помещений")) {
      ASSETS_ROW_NUM = rowIndex;
    }
  }
}

function traverseArrayAndFindRows(array) {
  array.forEach((row, rowIndex) => {
    if (row.length > 0) {
      row.forEach((rowValue, deepRowIndex) => {
        getTotalCountRowNum(rowValue, rowIndex);
        getAssetsRowNum(rowValue, rowIndex);
      });
    }
  });
}

/******* UI METHODS AND IMPLEMENTATIONS ***********/

function setChosenFiles(name) {
  const chosenFiles = document.querySelector(".chosenFiles");
  chosenFiles.textContent = `Выбранные файлы: ${name}`
}

var dropArea = document.querySelector(".input-area");
var label = document.querySelector("label");

dropArea.addEventListener('dragenter', () => {
  dropArea.style.background = 'rgba(255, 255, 255, 0.2)'
})