const xlsx = require('node-xlsx')
const fs = require('fs')
const docx = require('docx')

const input = document.querySelector('#fileInput')

input.addEventListener('change', () => {
  const file = input.files[0]
  const { path, name } = file
  setChosenFiles(name)
  const spreadsheetData = getSpreadsheetData({ spreadsheetPath: path })

  traverseArrayAndFindRows(spreadsheetData)
  const departmentName = spreadsheetData[1][1]

  const coreName = spreadsheetData[6][1]
  const assets = getAssets({ spreadsheetData })
  const assetsCount = getAssetsNum({ spreadsheetData })
  const assetsPositionsThatExist = getAssetsPositionThatExist({ assetsCount })
  logAssetsThatExist({ assetsPositionsThatExist, assets })

  const assetsInString = arrayAssetsThatExist({
    assetsPositionsThatExist,
    assets,
  }).join(',')
  const resultString = `${coreName},${assetsInString}`
  createdoc(name, departmentName, resultString)
})

/* ******** HELPERS METHODS ********************* */
var TOTAL_COUNT_ROW_NUM = 0
var ASSETS_ROW_NUM = 0
const DOCX_RESULTS_FOLDER_NAME = 'results'

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
  const out = []
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
      row.forEach(rowValue => {
        getTotalCountRowNum(rowValue, rowIndex)
        getAssetsRowNum(rowValue, rowIndex)
      })
    }
  })
}

// noinspection UnterminatedStatementJS
function createdoc(name, object, inventory) {
  // style example
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
    .font('Times New Roman')  
  .size(16)
    .basedOn('Normal')
    // .italics()
    .justified()
    .spacing({ line: 240, before: 0, after: 0 })
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
  doc.Styles.createParagraphStyle('itemlist', 'ItemList')
    .font('Times New Roman')
    .basedOn('Normal')
    .next('Normal')
    .quickFormat()
    .size(12)
    .underline()
    .spacing({ line: 240, before: 0, after: 0 })
  doc.Styles.createParagraphStyle('beforeitemlist', 'BeforeItemList')
    .font('Times New Roman')
    .basedOn('Normal')
    .next('Normal')
    .quickFormat()
    .size(24)
    .spacing({ line: 240, before: 0, after: 0 })
  doc.createParagraph(
    `АКТ \n \n`
    /* "технического освидетельствования средств и систем охраны" */
  )
    .style('Heading1')
    .center()
  //doc.addParagraph(paragraph)
  doc.createParagraph(
    'технического освидетельствования средств и систем охраны '
  )
    .style('Heading1')
    .center()
 // doc.addParagraph(paragraph)
  doc.createParagraph('охранно- тревожной сигнализации ')
    .style('UnderLine')
    .center()
 // doc.addParagraph(paragraph)

  doc.createParagraph(
    'наименование технических средств и систем охраны '
  )
    .style('UnderText')
    .center()
 // doc.addParagraph(paragraph)
 // doc.addParagraph(paragraph);
  doc.createParagraph("\nг.Минск \n\n").style('Heading1').justified();
 // doc.addParagraph(paragraph);
  doc.createParagraph("Комиссия в составе:\n " ).style('Heading1').left();
  doc.createParagraph(
    "-  ВрИОД инспектора - инженера отделения средств и систем охраны Партизанского (г.Минска) отдела Департамента охраны МВД Республики Беларусь Жука В.П.\n"  +
    "-  электромонтера охранно - пожарной сигнализации Партизанского (г.Минска) отдела Департамента охраны МВД Республики Беларусь ____________________ \n").style('Heading1').left();
  //doc.addParagraph(paragraph);
  var paragraph = new docx.Paragraph("произвела техническое освидетельствование  ").style('Heading1').left();
  var text ;
  //new docx.TextRun("My awesome text here for my university dissertation");
  if (typeof inventory === 'string') text = new docx.TextRun(inventory).style('itemlist').size(14).underline();
  
 // var text = new docx.TextRun("My awesome text here for my university dissertation");
  paragraph.addRun(text);
  
  doc.addParagraph(paragraph.justified());
  doc.createParagraph("наименование технических средств и систем охраны \n ").style('UnderText').right();
  if (typeof object === 'string')doc.createParagraph(object).style('itemlist').justified();
  //TODO: insert a lot of spaces
  doc.createParagraph("наименование объекта, жилого дома (помещения) физического лица, адрес, на котором они смонтированы  ").style('UnderText').center();
  doc.createParagraph("___________________________________________________________________________________ \n").style('Heading1').justified();

  doc.createParagraph("При техническом освидетельствовании установлено:\n    -существующие средства и системы охраны на момент проверки работоспособны во всех режимах;\n\n Комиссия рекомендует:\n  - допустить средства и системы охраны к дальнейшей эксплуатации в течение одного года до проведения следующего технического освидетельствования \n").style('Heading1').left();


  doc.createParagraph("\nЧлены комиссии:                                               ______________________________").style('Heading1').justified();
  doc.createParagraph("                                                                                                                        подпись, фамилия, инициалы  \n").style('UnderText').left();
  doc.createParagraph("                                                                             ______________________________").style('Heading1').left();
  doc.createParagraph("                                                                                                                        подпись, фамилия, инициалы  \n").style('UnderText').left();
  doc.createParagraph("                                                                             ______________________________").style('Heading1').left();
  doc.createParagraph("                                                                                                                        подпись, фамилия, инициалы  \n").style('UnderText').left();


  var packer = new docx.Packer()
  var newName = 'file'
  if (typeof name === 'string') newName = name.split('.')[0]

  const isResultsFolderExist = checkFolderExist({
    path: DOCX_RESULTS_FOLDER_NAME,
  })


  if (!isResultsFolderExist) 
    fs.mkdirSync(`./${DOCX_RESULTS_FOLDER_NAME}`)/*{
    return writeDocxFile({
      doc,
      path: DOCX_RESULTS_FOLDER_NAME,
      packer,
      name: newName,
    })
  }
  fs.mkdirSync(`./${DOCX_RESULTS_FOLDER_NAME}`)*/
  return writeDocxFile({
    doc,
    path: DOCX_RESULTS_FOLDER_NAME,
    packer,
    name: newName,
  })
}

function writeDocxFile({ packer, path, name, doc }) {
  packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(`./${path}/${name}.docx`, buffer)
  })
}

function checkFolderExist({ path }) {
  return fs.existsSync(`./${path}`)
}

/* ***** UI METHODS AND IMPLEMENTATIONS ********** */

function setChosenFiles(name) {
  const chosenFiles = document.querySelector('.chosenFiles')
  chosenFiles.textContent = `Выбранный файл: ${name}`
}

var dropArea = document.querySelector('.input-area')

dropArea.addEventListener('dragenter', () => {
  dropArea.style.background = 'rgba(255, 255, 255, 0.2)'
})
