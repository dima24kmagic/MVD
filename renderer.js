const xlsx = require("node-xlsx");
const fs = require("fs");

const input = document.querySelector("#fileInput");

let TOTAL_COUNT_ROW_NUM = 0;

const getSpreadsheetData = ({ spreadsheetPath }) => {
  const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(spreadsheetPath));
  const { data } = workSheetsFromBuffer[0];
  return data;
};

const getAssets = ({ spreadsheetData }) => {
  const assets = spreadsheetData[15];
  const filteredAssets = assets.filter(asset => asset !== undefined);
  return filteredAssets;
};

const getAssetsNum = ({ spreadsheetData }) => {
  const assetsTotalCount = spreadsheetData[81];
  return assetsTotalCount.filter(
    assetTotal => assetTotal !== undefined && assetTotal !== "Всего"
  );
};

const getAssetsPositionThatExist = ({ assetsCount }) => {
  const assetsExistPositions = [];
  assetsCount.forEach((asset, index) => {
    if (asset !== 0) {
      assetsExistPositions.push(index);
    }
  });
  return assetsExistPositions;
};

const logAssetsThatExist = ({ assetsPositionsThatExist, assets }) => {
  assetsPositionsThatExist.forEach(position => console.log(assets[position]));
};

const getTotalCountRowNum = (rowValue, rowIndex) => {
  if (typeof rowValue === "string") {
    if (rowValue.includes("всего")) {
      TOTAL_COUNT_ROW_NUM = rowIndex;
    }
  }
};

const traverseArray = array => {
  array.forEach((row, rowIndex) => {
    if (row.length > 0) {
      row.forEach((rowValue, deepRowIndex) => {
        getTotalCountRowNum(rowValue, rowIndex);
        // console.log(rowValue, "|", rowIndex + 1, deepRowIndex + 1);
      });
    }
  });
};

input.addEventListener("change", () => {
  const file = input.files[0];
  console.log({ file });
  const { path } = file;
});

const spreadsheetData = getSpreadsheetData({
  spreadsheetPath:
    "/home/dima/programming/projects/MVD/assets/УУ гроссоптик жасминовая 2г.xlsx"
});

traverseArray(spreadsheetData);

const departmentName = spreadsheetData[1][1];
const assets = getAssets({ spreadsheetData });
const assetsCount = getAssetsNum({ spreadsheetData });
const assetsPositionsThatExist = getAssetsPositionThatExist({ assetsCount });
logAssetsThatExist({ assetsPositionsThatExist, assets });