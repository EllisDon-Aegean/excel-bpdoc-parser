const xlsx = require('xlsx');
const jsonfile = require('jsonfile')

/* Index to map excel columns to specific document properties */
let docPropertiesIndex = require('./docPropertiesIndex.js');

/* Base path for GCP Bucket */
let GCP_FILEPATH = '/bestpractices-docs';
/* ALL DOCS WILL BE CATEGORIZED UNDER THE CATEGORY YOU INPUT.  */
let DOCS_CATEGORY = '';

/* SPECIFY THE EXCEL DOC TO PARSE HERE */
let path = (__dirname + '/assets/MEIT.xlsx');

/* Parsing the excel file. An excel is a 'workbook' and it can have multile 'sheets'. */
let workbook = xlsx.readFile(path);
var worksheet = Object.values(workbook.Sheets)[0]; /* In this case, were always using the first one. */

/* This script loops through the rows in an excel table. So for example,
it'll start in row 1, and parse the data in each column in that row untill
it reaches the 'lastRow' that you will specify below. When it does, it'll
move to the next row untill it reaches the last row that you'll specify below. */


/* --- CHANGE THESE VALUES DEPENDING ON YOUR EXCEL FILE --- */

/*
In the excel file, where the data starts (exclude the column title.)
ie. firstColumn = 'A' lastColumn = 1 || 'A1' is where the data begins.
*/
let firstColumn = 'B';
let firstRow = 2;
/* Where the data ends.. (ie. 'F10') */
let lastColumn = 'H';
let lastRow = 29;

/* --------------------------------- */

firstColumn = firstColumn.charCodeAt(0);
lastColumn = lastColumn.charCodeAt(0);

/* --- HELPER FUNCTIONS -- */

function addPropertiesAndValueToDoc(row, column, doc) {
  if (column !== lastColumn + 1) {
    let tags = ['phase', 'type', 'document_link', 'gcp_filepath'];
    let currentCell = String.fromCharCode(column) + `${row}`; /* ie. 'A8' */
    let documentProperty = docPropertiesIndex[currentCell[0]];
    let value = worksheet[currentCell] ? worksheet[currentCell].v : ''

    /* If its a tag. */
    if (tags.includes(documentProperty)) {

      if (documentProperty == 'phase') {
        doc.tags[documentProperty] = value ? determineApplicablePhases(value) : '';
      } else if (documentProperty == 'gcp_filepath') {
        doc.tags[documentProperty] = value ? determineGCPFilePath(value, doc.section) : '';
      }  else {
        doc.tags[documentProperty] = value
      }

    } else {
      doc[documentProperty] = value;
    }

    addPropertiesAndValueToDoc(row, column + 1, doc);
  }
  return doc;
}

function createDocumentsArray(row, column, docsArray) {
  if (row !== lastRow + 1) {
    doc = addPropertiesAndValueToDoc(row, column, {tags: {}, categories: [DOCS_CATEGORY]});
    docsArray.push(doc);
    createDocumentsArray(row + 1, column, docsArray);
  }
  return docsArray;
}

function determineApplicablePhases(phasesString) {
  /* sometimes a number */
  let applicablePhasesByNumber = phasesString.toString().replace(/,/g, '').replace(" ", '').split(''); /* ie. [1, 2, 4] */
  let phasesIndex = ['Project Pursuit', 'Project Planning', 'Project Execution', 'Warranty / Facility Management'];
  let applicablePhases = applicablePhasesByNumber.map((phaseNumber) => {
    return phasesIndex[eval(phaseNumber - 1)]
  })

  return applicablePhases;
}

function determineGCPFilePath(fileName, section) {
  return `${GCP_FILEPATH}/${section.toLowerCase()}/${fileName}`;
}

/* --- --- */

let docsObject = {
  docs: createDocumentsArray(firstRow, firstColumn, [])
}

jsonfile.writeFile('data.json', docsObject, { spaces: 2 })
.then(() => console.log('Writing to JSON file done.'))
.catch((e) => console.log('Error in writing to JSON file', e))
