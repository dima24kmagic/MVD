const xlsx = require('node-xlsx')
const fs = require('fs')
const docx = require('docx')

const input = document.querySelector('#fileInput')

input.addEventListener('change', () => {
  const file = input.files[0]
  const { path, name } = file
  setChosenFiles(name)
  const spreadsheetData = getSpreadsheetData({ spreadsheetPath: path })
  var resultString = ''

  traverseArrayAndFindRows(spreadsheetData)
  const departmentName = spreadsheetData[1][1]

  const coreName = spreadsheetData[6][1]
  const assets = getAssets({ spreadsheetData })
  const assetsCount = getAssetsNum({ spreadsheetData })
  const assetsPositionsThatExist = getAssetsPositionThatExist({ assetsCount })
  logAssetsThatExist({ assetsPositionsThatExist, assets })

  resultString +=
    departmentName +
    ',' +
    coreName +
    ',' +
    arrayAssetsThatExist({ assetsPositionsThatExist, assets }).join(',')
  createdoc(name, resultString)
})

/********* HELPERS METHODS **********************/

var TOTAL_COUNT_ROW_NUM = 0
var ASSETS_ROW_NUM = 0

function getSpreadsheetData({ spreadsheetPath }) {
  const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(spreadsheetPath))
  const { data } = workSheetsFromBuffer[0]
  return data
}

function getAssets({ spreadsheetData }) {
  const assets = spreadsheetData[ASSETS_ROW_NUM + 1]
  const filteredAssets = assets.filter(asset => asset !== undefined)
  return filteredAssets
}

function getAssetsNum({ spreadsheetData }) {
  const assetsTotalCount = spreadsheetData[TOTAL_COUNT_ROW_NUM]
  return assetsTotalCount.filter(
    assetTotal => assetTotal !== undefined && assetTotal !== 'Всего'
  )
}

function getAssetsPositionThatExist({ assetsCount }) {
  const assetsExistPositions = []
  assetsCount.forEach((asset, index) => {
    if (asset !== 0) {
      assetsExistPositions.push(index)
    }
  })
  return assetsExistPositions
}

function logAssetsThatExist({ assetsPositionsThatExist, assets }) {
  assetsPositionsThatExist.forEach(position => console.log(assets[position]))
}
function arrayAssetsThatExist({ assetsPositionsThatExist, assets }) {
  var out = new Array()
  assetsPositionsThatExist.forEach(position => out.push(assets[position]))
  return out
}

function getTotalCountRowNum(rowValue, rowIndex) {
  if (typeof rowValue === 'string') {
    if (rowValue.search(/всего/i) !== -1) {
      TOTAL_COUNT_ROW_NUM = rowIndex
    }
  }
}
function getAssetsRowNum(rowValue, rowIndex) {
  if (typeof rowValue === 'string') {
    if (rowValue.search(/помещений/i) !== -1) {
      ASSETS_ROW_NUM = rowIndex
    }
  }
}

function traverseArrayAndFindRows(array) {
  array.forEach((row, rowIndex) => {
    if (row.length > 0) {
      row.forEach((rowValue, deepRowIndex) => {
        getTotalCountRowNum(rowValue, rowIndex)
        getAssetsRowNum(rowValue, rowIndex)
      })
    }
  })
}

function createdoc(name, text) {
  //style example
  var doc = new docx.Document(undefined, {
    top: 0,
    right: 556,
    bottom: 0,
    left: 1250,
  })
  doc.Styles.createParagraphStyle('wellSpaced', 'Well Spaced')
    .basedOn('Normal')
    .color('999999')
    .italics()
    .spacing({ line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 })
  doc.Styles.createParagraphStyle('default', 'Default')
    .basedOn('Normal')
    .color('999999')
    .size(24)
    // .italics()
    .justified()
    .spacing({ line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 })
  doc.Styles.createParagraphStyle('underrext', 'UnderText')
    .size(16)
    .basedOn('Normal')
    // .italics()
    .justified()
    .spacing({ line: 240, before: 20, after: 20 })
  doc.Styles.createParagraphStyle('Heading1', 'Heading 1')
    .font('Times New Roman')
    .basedOn('Normal')
    .next('Normal')
    .quickFormat()
    .size(24)
    .spacing({ line: 240, before: 20, after: 20 })
  doc.Styles.createParagraphStyle('underline', 'UnderLine')
    .font('Times New Roman')
    .basedOn('Normal')
    .next('Normal')
    .quickFormat()
    .size(24)
    .underline()
    .spacing({ line: 240, before: 20, after: 20 })
  var paragraph = new docx.Paragraph(
    'АКТ' +
      '\n' +
      '\n' /*
  "технического освидетельствования средств и систем охраны"*/
  )
    .style('Heading1')
    .center()
  doc.addParagraph(paragraph)
  paragraph = new docx.Paragraph(
    'технического освидетельствования средств и систем охраны '
  )
    .style('Heading1')
    .center()
  doc.addParagraph(paragraph)
  paragraph = new docx.Paragraph('охранно- тревожной сигнализации ')
    .style('UnderLine')
    .center()
  doc.addParagraph(paragraph)

  paragraph = new docx.Paragraph(
    'наименование технических средств и систем охраны '
  )
    .style('UnderText')
    .center()
  doc.addParagraph(paragraph)
  if (typeof text === 'string') doc.addParagraph(new docx.Paragraph(text))
  var packer = new docx.Packer()
  var newName = 'file'
  if (typeof name === 'string') newName = name.split('.')[0]

  // TODO: Check that folder "results" exist
  packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync('./results/' + newName + '.docx', buffer)
  })
}

/******* UI METHODS AND IMPLEMENTATIONS ***********/

function setChosenFiles(name) {
  const chosenFiles = document.querySelector('.chosenFiles')
  chosenFiles.textContent = `Выбранный файл: ${name}`
}

var dropArea = document.querySelector('.input-area')
var label = document.querySelector('label')

dropArea.addEventListener('dragenter', () => {
  dropArea.style.background = 'rgba(255, 255, 255, 0.2)'
})
