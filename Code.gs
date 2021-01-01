// Change "false" to "true" to see more detailed diagnostic information in the logs
var DEBUG = false;

function merge() {
  const inputSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const inputSheet = SpreadsheetApp.getActiveSheet()
  
  const headerRow = inputSheet.getRange(1, 1, 1, inputSheet.getLastColumn())
  if (DEBUG) {
    Logger.log("------Sheet header------")
    debugSheet(headerRow)
  }
  
  const dataRows = inputSheet.getRange(2, 1, inputSheet.getLastRow() - 1, inputSheet.getLastColumn())
  if (DEBUG) {
    Logger.log("------Sheet data------")
    debugSheet(dataRows)
  }
  
  const iter = DriveApp.getFilesByName(inputSheet.getName())
  let firstFile
  if (!iter.hasNext()) {
    // No files matching name
    Logger.log("No files matching Sheet name `" + inputSheet.getName() + "'. Skipping")
    return
  }
  // First file matching name
  firstFile = iter.next()
  // If there are more than one files matching the name, skip out
  if (iter.hasNext()) {
    Logger.log("More than 1 file matching Sheet name '" + inputSheet.getName() + "'. Skipping")
    return
  }
  
  const templateDoc = DocumentApp.openByUrl(firstFile.getUrl())
  
  const templateBody = templateDoc.getBody()
  if (DEBUG) {
    Logger.log("------Template------")
    debugElement(0, "Root", templateBody)
  }

  // Create output document
  const outputDocumentName = getOutputDocumentName(templateDoc.getName())
  if (DEBUG) {
    Logger.log("Output document name = " + outputDocumentName)
  }

  const outputDoc = DocumentApp.create(outputDocumentName)
  let outputBody = outputDoc.getBody()
  if (DEBUG) {
    Logger.log("outputDoc = " + outputDoc)
    Logger.log("outputDoc.getId = " + outputDoc.getId())
  }
  
  startMerge(templateBody, headerRow, dataRows, outputBody)
  
  const response = SpreadsheetApp.getUi().alert("Documents merged, and output is available in the document named \"" + outputDocumentName + "\".")
}

function startMerge(templateBody, headerRow, dataRows, outputBody) {
  const variableNames = headerRow.getValues()[0]
  const dataRowValues = dataRows.getValues()
  if (DEBUG) {
    Logger.log("dataRowValues = " + dataRowValues)
  }
  
  if ((templateBody.getNumChildren() == 3) && (templateBody.getChild(0).getType() == DocumentApp.ElementType.PARAGRAPH) && (templateBody.getChild(2).getType() == DocumentApp.ElementType.PARAGRAPH) && (templateBody.getChild(0).asParagraph().getText().length == 0) && (templateBody.getChild(2).asParagraph().getText().length == 0) && (templateBody.getChild(1).getType() == DocumentApp.ElementType.TABLE)) {
    Logger.log("Special condition")
    const inputTable = templateBody.getChild(1).asTable()
    const outputTable = outputBody.appendTable(createFormattedEmptyCopy(inputTable))
    for (var dataRow in dataRowValues) {
      performMerge(inputTable, variableNames, dataRowValues[dataRow], outputTable)
    }
  } else {
    for (var dataRow in dataRowValues) {
      performMerge(templateBody, variableNames, dataRowValues[dataRow], outputBody)
      outputBody.appendPageBreak()
    }
  }
}

function performMerge(templateElement, variableNames, dataRow, outputElement) {
  const numChildren = templateElement.getNumChildren()
  for (let templateChildIndex = 0; templateChildIndex < numChildren; templateChildIndex++) {
    let templateChild = templateElement.getChild(templateChildIndex)
    const elementType = templateChild.getType()
    switch (elementType) {
      case DocumentApp.ElementType.PARAGRAPH:
        const templatePara = templateChild.asParagraph()
        const originalText = templatePara.getText()
        let paraToAppend = templatePara.copy()
        paraToAppend = replaceAllVariablesFromElement(paraToAppend, variableNames, dataRow)
        const appended = outputElement.appendParagraph(paraToAppend)
        break
      case DocumentApp.ElementType.TABLE:
        const templateTableCopy = createFormattedEmptyCopy(templateChild)
        let outputTable = outputElement.appendTable(templateTableCopy)
        if (DEBUG) {
          Logger.log("Got TABLE type, appended table in output and calling recursively")
        }
        performMerge(templateChild, variableNames, dataRow, outputTable)
        break
      case DocumentApp.ElementType.TABLE_ROW:
        const templateRowCopy = createFormattedEmptyCopy(templateChild)
        let outputRow = outputElement.appendTableRow(templateRowCopy)
        if (DEBUG) {
          Logger.log("Got TABLE_ROW type, appended table in output and calling recursively")
        }
        performMerge(templateChild, variableNames, dataRow, outputRow)
        break
      case DocumentApp.ElementType.TABLE_CELL:
        const templateCellCopy = createFormattedEmptyCopy(templateChild)
        let outputCell = outputElement.appendTableCell(templateCellCopy)
        if (DEBUG) {
          Logger.log("Got TABLE_CELL type, appended table in output and calling recursively")
        }
        performMerge(templateChild, variableNames, dataRow, outputCell)
        break
    }
  }
}

function createFormattedEmptyCopy(element) {
  let elementCopy = element.copy()
  for (var childIndex = element.getNumChildren() - 1; childIndex >= 0; childIndex--) {
    elementCopy.removeChild(elementCopy.getChild(childIndex))
  }
  return elementCopy
}
function replaceAllVariablesFromElement(element, variableNames, dataRow) {
  for (var variableIndex = 0; variableIndex < variableNames.length; variableIndex++) {
    let variableName = variableNames[variableIndex]
    element = element.replaceText("<<"+variableName+">>", dataRow[variableIndex])
  }
  return element
}

function debugElement(depth, prefix, element) {
  const numChildren = element.getNumChildren()
  Logger.log(depth + "-numChildren = " + numChildren)
  for (var i = 0; i < numChildren; i++) {
    let child = element.getChild(i)
    Logger.log(depth + "[" + i + "].child = " + child)
    const elementType = child.getType()
    switch (elementType) {
      case DocumentApp.ElementType.PARAGRAPH:
        Logger.log(depth + "-" + prefix + "-Paragraph '" + child.asParagraph().getText() + "'")
        break
      default:
        debugElement(depth + 1, prefix+"."+elementType, child)
        break
    }
  }
}

function debugSheet(rows) {
  const values = rows.getValues()
  for (let row in values) {
    for (let col in values[row]) {
      let logMessage = "row " + row + ", col " + col + " has value " + values[row][col]
      Logger.log(logMessage)
    }
  }
}

function getOutputDocumentName(inputDocumentName) {
  const currentTimestamp = new Date()
  const time = Utilities.formatDate(currentTimestamp, "GMT", "yyyyMMdd-HHmmss-SSS")
  return inputDocumentName + "-output-" + time
}

function onInstall() {
  onOpen()
}

function onOpen() {
  const doc = DocumentApp.getActiveDocument()
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  
  let menuEntries = []
  menuEntries.push({name: "Merge", functionName: "merge"})
  ss.addMenu("Merge", menuEntries)
}