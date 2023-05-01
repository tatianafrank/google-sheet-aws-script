function updateReport() {
  const spreadsheetId = SPREADSHEET_ID;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let fbaSheet = ss.getSheetByName('Report-Data-Pull')
  let fbmSheet = ss.getSheetByName('FBM-Report')
  let yourDate = new Date()
  yourDate.toISOString().split('T')[0]
  const offset = yourDate.getTimezoneOffset()
  yourDate = new Date(yourDate.getTime() - (offset*60*1000))
  const todayDate = yourDate.toISOString().split('T')[0]
  // const fbaAsins = fbaSheet.getRange("E2:E").getValues().filter((cell) => cell[0].length !== 0);
  // const fbmAsins = fbmSheet.getRange("C2:C").getValues().filter((cell) => cell[0].length !== 0);
  const fbaFileName = 'fba_' + todayDate.toString();
  const fbmFileName = 'fbm_' + todayDate.toString();

  if (getTodayDriveReport(fbaFileName) === false) {
    Logger.log('Creating new FBA report to upload to drive')
    return false;
  }; // if report is being created stop script, try again in 10 mins

  if (getTodayDriveReport(fbmFileName) === false) {
    Logger.log('Creating new FBM report to upload to drive')
    return false;
  }; // if report is being created stop script, try again in 10 mins

  checkForUpdates(fbaSheet, fbaFileName, todayDate, 'fba') // Line by line update
  checkForUpdates(fbmSheet, fbmFileName, todayDate, 'fbm')

}

function checkForUpdates(theSheet, fileName, todayDate, ftype) {
    const allDateValues = theSheet.getRange('A2:A').getValues();
    let firstRowOutdated = -1;
    const twoDaysInMs = 2 * 24 * 60 * 60 * 1000;
    const timestampTwoDaysAgo = new Date().getTime() - twoDaysInMs;
    Logger.log(['Updating ', ftype, ' report']);

    if (!(allDateValues[0][0])) {
      Logger.log(['No first data value so uploading todays report to sheets']) 
      const reportBlob = getReport(fileName);
      uploadReportToSheets(reportBlob, theSheet, todayDate, 0, ftype)
    } else {
      for (i=0; i< allDateValues.length; i++) {
        if (allDateValues[i][0]) {
          try {
            ddate = new Date(allDateValues[i][0]);
            if (timestampTwoDaysAgo> ddate) {
              firstRowOutdated = i+1;
              Logger.log(['First row older than two days is: ', firstRowOutdated, ' Uploading report starting at row ', firstRowOutdated]) 
              const reportBlob = getReport(fileName);
              uploadReportToSheets(reportBlob, theSheet, todayDate, firstRowOutdated, ftype)
              break;
            }
          } catch (error) {
            Logger.log(error)
          }
        }
      }
    }

}
function addCustomFbaColumns(reportSheet) {
  if (reportSheet.getRange("G1").getValues()[0][0] !== "Product Image Url") {
      Logger.log('adding new columns')
      reportSheet.insertColumnAfter(6);
      reportSheet.getRange('G1').setValue('Product Image Url');
      reportSheet.insertColumnAfter(7);
      reportSheet.getRange('H1').setValue('Product Image');
      reportSheet.insertColumnAfter(8);
      reportSheet.getRange('I1').setValue('1st most popular category id');
      reportSheet.insertColumnAfter(9);
      reportSheet.getRange('J1').setValue('1st most popular category rank');
      reportSheet.insertColumnAfter(10);
      reportSheet.getRange('K1').setValue('2nd most popular category id');
      reportSheet.insertColumnAfter(11);
      reportSheet.getRange('L1').setValue('2nd most popular category rank');
      reportSheet.insertColumnAfter(12);
      reportSheet.getRange('M1').setValue('3rd most popular category id');
      reportSheet.insertColumnAfter(13);
      reportSheet.getRange('N1').setValue('3rd most popular category rank');
    }

    // const numRows = Math.floor(asins.length+1);
    // let imgRows = reportSheet.getRange("G1:G").getValues();
    // let firstEmptyRow = Math.floor(imgRows.filter((cell) => cell[0].length !== 0).length +1);
    // Logger.log(['first empty image row: ', firstEmptyRow])
    // for (let indx = firstEmptyRow; indx < numRows+1; indx++) {
    //   // start filling in the img/rank rows starting from the first empty cell
    //   // if none are empty, loop wont run
    //   Logger.log(['asins length ', asins.length, 'indx -2 ', indx-2, asins[asins.length-1]])
    //   asin = asins[indx-2][0];
    //   let imagesAndRanks = getImagesAndRanks(asin);
    //   uploadImagesRanksToSheets(JSON.parse(imagesAndRanks)[asin], reportSheet, indx);
    // }
}

function addCustomFbmColumns(fbmSheet) {
    if (fbmSheet.getRange("G1").getValues()[0][0] !== "Product Image URL") {
      fbmSheet.insertColumnAfter(6);
      fbmSheet.getRange('G1').setValue('Product Image URL');
      fbmSheet.insertColumnAfter(7);
      fbmSheet.getRange('H1').setValue('Product Image');
      fbmSheet.insertColumnAfter(8);
      fbmSheet.getRange('I1').setValue('7 Day Order Count');
      fbmSheet.insertColumnAfter(9);
      fbmSheet.getRange('J1').setValue('7 Day Average Unit Price');
      fbmSheet.insertColumnAfter(10);
      fbmSheet.getRange('K1').setValue('30 Day Order Count');
      fbmSheet.insertColumnAfter(11);
      fbmSheet.getRange('L1').setValue('30 Day Average Unit Price');
      fbmSheet.insertColumnAfter(12);
      fbmSheet.getRange('M1').setValue('60 Day Order Count');
      fbmSheet.insertColumnAfter(13);
      fbmSheet.getRange('N1').setValue('60 Day Average Unit Price');
      fbmSheet.insertColumnAfter(14);
      fbmSheet.getRange('O1').setValue('90 Day Order Count');
      fbmSheet.insertColumnAfter(15);
      fbmSheet.getRange('P1').setValue('90 Day Average Unit Price');
      fbmSheet.insertColumnAfter(16);
      fbmSheet.getRange('Q1').setValue('Buy Box Listing Price');
      fbmSheet.insertColumnAfter(17);
      fbmSheet.getRange('R1').setValue('Lowest Price Listing Price');
    }

  // const numRows = Math.floor(asins.length+1);
  // let salesRows = fbmSheet.getRange("I1:I").getValues();
  // let firstEmptyRow = Math.floor(salesRows.filter((cell) => cell[0].length !== 0).length +1);
  // for (let indx = firstEmptyRow; indx < numRows+1; indx++) {
  //   let asin = asins[indx-2][0];
  //   sales = getSales(asin);
  //   uploadSalesToSheets(sales, fbmSheet, indx);
  //   try {
  //     let imgUrl = getImageUrl(asin);
  //     uploadImageToFbm(imgUrl, fbmSheet, indx);
  //     let pricing = getPrices(asin);
  //     uploadPricingToFbm(pricing, fbmSheet, indx)
  //   } catch(err) {
  //     let theRange = indx+':'+indx;
  //     fbmSheet.getRange(theRange).setBackground("red");
  //   }
  // }
}

function getPrices(asin) {
    const options = {
    "method": "get",
    "contentType": 'application/json',
    "payload": JSON.stringify({
      "auth": AUTH_NUM,
      "endpoint": "get_fbm_pricing",
      "asins": asin
    }),
    // muteHttpExceptions: true
  };
  return callLambda(options);
}

function uploadPricingToFbm(pricing, fbmSheet, indx) {
  pricing = JSON.parse(pricing)
  if ('BuyBoxPrices' in pricing['Summary']) {
    const priceCellQ = 'Q'+ indx.toString();
    const buyBoxPrice = pricing['Summary']['BuyBoxPrices'][0]['ListingPrice']['Amount']
    fbmSheet.getRange(priceCellQ).setValue(JSON.stringify(buyBoxPrice));
  }
  if ('LowestPrices' in pricing['Summary']) {
    const lowestPrice = pricing['Summary']['LowestPrices'][0]['ListingPrice']['Amount']
    const priceCellR = 'R'+ indx.toString();
    fbmSheet.getRange(priceCellR).setValue(JSON.stringify(lowestPrice));
  }
}

function uploadImageToFbm(imgUrl, fbmSheet, indx) {
  const imgCellG = 'G'+ indx.toString();
  const imgCellH = 'H'+ indx.toString();
  fbmSheet.getRange(imgCellG).setValue(imgUrl);
  try {
    fbmSheet.getRange(imgCellH).setFormula(`=IMAGE(${imgCellG})`);
    SpreadsheetApp.flush();
  } catch(err) {
    Logger.log(err);
  }
}

function getImageUrl(asin){
  const options = {
    "method": "get",
    "contentType": 'application/json',
    "payload": JSON.stringify({
      "auth": AUTH_NUM,
      "endpoint": "get_image_url",
      "asins": asin
    }),
    muteHttpExceptions: true
  };
  return callLambda(options);
}


function uploadSalesToSheets(sales, fbmSheet, indx) {
  let saleCellI = 'I'+ indx.toString();
  let saleCellJ = 'J'+ indx.toString();
  let saleCellK = 'K'+ indx.toString();
  let saleCellL = 'L'+ indx.toString();
  let saleCellM = 'M'+ indx.toString();
  let saleCellN = 'N'+ indx.toString();
  let saleCellO = 'O'+ indx.toString();
  let saleCellP = 'P'+ indx.toString();

  sales = JSON.parse(sales)
  fbmSheet.getRange(saleCellI).setValue(JSON.stringify(sales["week"]["orderCount"]));
  fbmSheet.getRange(saleCellJ).setValue(JSON.stringify(sales["week"]["averageUnitPrice"]["amount"]));
  fbmSheet.getRange(saleCellK).setValue(JSON.stringify(sales["one_month"]["orderCount"]));
  fbmSheet.getRange(saleCellL).setValue(JSON.stringify(sales["one_month"]["averageUnitPrice"]["amount"]));
  fbmSheet.getRange(saleCellM).setValue(JSON.stringify(sales["two_month"]["orderCount"]));
  fbmSheet.getRange(saleCellN).setValue(JSON.stringify(sales["two_month"]["averageUnitPrice"]["amount"]));
  fbmSheet.getRange(saleCellO).setValue(JSON.stringify(sales["three_month"]["orderCount"]));
  fbmSheet.getRange(saleCellP).setValue(JSON.stringify(sales["three_month"]["averageUnitPrice"]["amount"]));
}

function getSales(asin) {
    const options = {
      "method": "get",
      "contentType": 'application/json',
      "payload": JSON.stringify({
        "auth": AUTH_NUM,
        "endpoint": "get_sales",
        "asins": asin
      }),
      // muteHttpExceptions: true
    };
    return callLambda(options);
}
function getTodayDriveReport(fileName) {
  const parentFolder = DriveApp.getFolderById(FOLDER_ID);
  const files = parentFolder.getFilesByName(fileName)
  if (!files.hasNext()) {
    Logger.log(['creating report... will try to upload to sheets in an hour'])
    createReport(fileName);
    return false;
  } else {
    return true;
  }
}

function getImagesAndRanks(asin){ 
  //  Print the values from spreadsheet if values are available.
  if (!asin) {
    Logger.log('No asin provided.');
    return 'no asin provided';
  }
  const options = {
    "method": "get",
    "contentType": 'application/json',
    "payload": JSON.stringify({
      "auth": AUTH_NUM,
      "endpoint": "get_images",
      "asins": asin
    }),
    // muteHttpExceptions: true
  };
  return callLambda(options);
  
}

function uploadReportToSheets(reportBlob, theSheet, filename, firstRowOutdated, ftype) {
  let lines = reportBlob.split('\n');
  for (let i = firstRowOutdated; i < (lines.length-1); i++) {
    let line = '';
    if(i==0) {
      line = 'Created Report Date' + '\t' + lines[i];
    } else {
      line = filename + '\t' + lines[i];
    }
    let row = line.split('\t');
    let ii = (i+1).toString();
    let theRange = ii+':'+ii;
    let oldRow = theSheet.getRange(theRange).getValues();
    updateRow(oldRow, row, theSheet, ii, ftype);
    
  }
}

function updateRow(oldRow, newRow, theSheet, rowNum, ftype) {
  if (rowNum == 1) {
    theSheet.getRange(rowNum, 1, 1, newRow.length).setValues([newRow]);
    if (ftype == 'fba') {
      addCustomFbaColumns(theSheet)
    }
    if (ftype == 'fbm') {
      addCustomFbmColumns(theSheet)
    }
  }
  if (rowNum > 1) {
    // let hasChanged = 0;
    if (ftype == 'fba') {
      theSheet.getRange(rowNum, 1, 1, 6).setValues([newRow.slice(0,6)])
      theSheet.getRange(rowNum, 15, 1, (newRow.length-6)).setValues([newRow.slice(6)]);
      let asin = newRow[4]
      Logger.log(['asin: ', asin])
      let imagesAndRanks = getImagesAndRanks(asin);
      uploadImagesRanksToSheets(JSON.parse(imagesAndRanks)[asin], theSheet, rowNum);
    }
    if (ftype == 'fbm') {
      // 7-18 = cust cols
      // getRange(row, column, numRows, numColumns) 
      theSheet.getRange(rowNum, 1, 1, 6).setValues([newRow.slice(0,6)])
      theSheet.getRange(rowNum, 19, 1, (newRow.length-6)).setValues([newRow.slice(6)]);
      let asin = newRow[2];
      Logger.log(['new row: ', newRow])
      Logger.log(['asin', asin])
      sales = getSales(asin);
      uploadSalesToSheets(sales, theSheet, rowNum);
      try {
        let imgUrl = getImageUrl(asin);
        uploadImageToFbm(imgUrl, theSheet, rowNum);
        let pricing = getPrices(asin);
        uploadPricingToFbm(pricing, theSheet, rowNum)
        let theRange = rowNum+':'+rowNum;
        theSheet.getRange(theRange).setBackground("white");
      } catch(err) {
        let theRange = rowNum+':'+rowNum;
        fbmSheet.getRange(theRange).setBackground("red");
      }
    }

    // for (let i = 14; i < oldRow[0].length; i++) {
    //   if(typeof(oldRow[0][i]) == 'number' && Number(oldRow[0][(i-8)]) !== Number(newRow[i])) {
    //     if(String(oldRow[0][i]).slice(0,10) !== String(newRow[(i-8)]).slice(0,10)) {
    //       Logger.log('old and new values are different')
    //       Logger.log(['number csv value: ', oldRow[0][i], typeof(oldRow[0][i]), 'report value: ', newRow[i], typeof(newRow[i])])
    //       hasChanged++;
    //     }
    //   } else if (typeof(oldRow[0][i]) == 'string' && oldRow[0][(i-8)] !== newRow[i]) {
    //     if (!(oldRow[0][i] == '' && newRow[(i-8)] == null)) {
    //       Logger.log('old and new values are different')
    //       Logger.log(['string old value: ', oldRow[0][i], typeof(oldRow[0][i]), 'new value: ', newRow[(i-8)], typeof(newRow[(i-8)])])
    //       hasChanged++;
    //     }
    //   }
    // }
    // if (hasChanged == 0) {
    //   Logger.log(['Nothing changed so only updating columns 1,2 for row: ', rowNum])
    //   theSheet.getRange(rowNum, 1, 1, 2).setValues([[newRow[0], newRow[1]]])
    // }
    // if (hasChanged > 0) {
    //   Logger.log(['Something changed'])
    // }
  }
}

function uploadImagesRanksToSheets(imagesRanks, sheet, indx) {

  let urlCell = 'G'+ indx.toString();
  let imgCell = 'H'+ indx.toString();
  let rankCellI = 'I'+ indx.toString();
  let rankCellJ = 'J'+ indx.toString();
  let rankCellK = 'K'+ indx.toString();
  let rankCellL = 'L'+ indx.toString();
  let rankCellM = 'M'+ indx.toString();
  let rankCellN = 'N'+ indx.toString();

  // [{"ProductCategoryId":"fashion_display_on_website","Rank":4487746},{"ProductCategoryId":"6358540011","Rank":29880},{"ProductCategoryId":"fashion_display_on_website","Rank":4487743},{"ProductCategoryId":"6358540011","Rank":29880}]
  let ranks = imagesRanks["rankings"];
  if(ranks.length > 0 && typeof(ranks) !== 'string') {
    if(ranks.length > 1) {
      let uniqueRankStrings = new Set();
      ranks = ranks.filter(element => {
        const isDuplicate = uniqueRankStrings.has(JSON.stringify(element));

        uniqueRankStrings.add(JSON.stringify(element));

        if (!isDuplicate) {
          return true;
        }
        return false;
      });
      ranks.sort(function (a,b) {
        return a['Rank'] - b['Rank'];
      });
      sheet.getRange(rankCellK).setValue(JSON.stringify(ranks[1]["ProductCategoryId"]));
      sheet.getRange(rankCellL).setValue(JSON.stringify(ranks[1]["Rank"]));
    }
    sheet.getRange(rankCellI).setValue(JSON.stringify(ranks[0]["ProductCategoryId"]));
    sheet.getRange(rankCellJ).setValue(JSON.stringify(ranks[0]["Rank"])); //if more than 1 this has to wait for above thats why its down here
    if(ranks.length > 2) {
      sheet.getRange(rankCellM).setValue(JSON.stringify(ranks[2]["ProductCategoryId"]));
      sheet.getRange(rankCellN).setValue(JSON.stringify(ranks[2]["Rank"]));
    }
  }

  sheet.getRange(urlCell).setValue(imagesRanks["imgUrl"]);
  // let blob = UrlFetchApp.fetch(imagesRanks["imgUrl"]).getBlob();
  // let base64String = Utilities.base64Encode(blob.getBytes());
  // let imageUrl =  `data:image/png;base64,${base64String}`;
  try {
    sheet.getRange(imgCell).setFormula(`=IMAGE(${urlCell})`);
    SpreadsheetApp.flush();
    // let image = SpreadsheetApp
    //               .newCellImage()
    //               .setSourceUrl(imageUrl)
    //               .setAltTextTitle('desc')
    //               .build()
    //               .toBuilder();
    // sheet.getRange(imgCell).setValue(image);
  } catch(err) {
    Logger.log(err);
  }
  
  
}

function getReport(fileName) {
  const folderId = FOLDER_ID;
  const folder = DriveApp.getFolderById(folderId);
  let files = folder.getFilesByName(name=fileName);
  if (!files.hasNext()) {
    return false;
  }
  while (files.hasNext()) {
    let fileContents = files.next().getBlob();
    return fileContents.getDataAsString();
  }
}

function createReport(fileName) {
  let endpoint = 'get_' + fileName.slice(0,3);
    const options = {
      "method": "get",
      "contentType": 'application/json',
      "payload": JSON.stringify({
        "auth": AUTH_NUM,
        "endpoint": endpoint,
      }),
      // muteHttpExceptions: true
    };
  let lambdaResponse = callLambda(options);
  const parentFolder = DriveApp.getFolderById(FOLDER_ID);
  
  if (fileName.slice(0,3) == 'fba') {
    const decoded = Utilities.base64Decode(lambdaResponse, Utilities.Charset.UTF_8)
    const decodedString = Utilities.newBlob(decoded).getDataAsString();
    return parentFolder.createFile(fileName, decodedString);
  } else if((fileName.slice(0,3) == 'fbm')) {
    const doc_url = lambdaResponse;
    const doc_options = {
      "method": "get",
      "contentType": 'application/json',
      "headers": {
        "Accept": "application/a-gzip"
      },
      // muteHttpExceptions: true
    };
    const charData = UrlFetchApp.fetch(doc_url, doc_options).getContent();
    eval(UrlFetchApp.fetch('https://cdn.rawgit.com/nodeca/pako/master/dist/pako.js').getContentText());
    const binData = [];
    for (let i = 0; i < charData.length; i++) {
      binData.push(charData[i] < 0 ? charData[i] + 256 : charData[i]);
    }

    let data = pako.inflate(binData);
    let decoded_chars = '';
    for (let i = 0; i < data.length; i++) {
      decoded_chars += String.fromCharCode(data[i]);
    }
    return parentFolder.createFile(fileName, decoded_chars);
  } 
}

function callLambda(options) {
  Logger.log('calling lambda');
  apiUrl = LAMBDA_URL;
  Logger.log(options)
  const lambdaResponse = UrlFetchApp.fetch(apiUrl, options).getContentText();
  Logger.log(['lambda response: ', lambdaResponse])
  return lambdaResponse;
}
,
        "endpoint": endpoint,
      }),
      // muteHttpExceptions: true
    };
  let lambdaResponse = callLambda(options);
  const parentFolder = DriveApp.getFolderById(FOLDER_ID);
  
  if (fileName.slice(0,3) == 'fba') {
    const decoded = Utilities.base64Decode(lambdaResponse, Utilities.Charset.UTF_8)
    const decodedString = Utilities.newBlob(decoded).getDataAsString();
    return parentFolder.createFile(fileName, decodedString);
  } else if((fileName.slice(0,3) == 'fbm')) {
    const doc_url = lambdaResponse;
    const doc_options = {
      "method": "get",
      "contentType": 'application/json',
      "headers": {
        "Accept": "application/a-gzip"
      },
      // muteHttpExceptions: true
    };
    const charData = UrlFetchApp.fetch(doc_url, doc_options).getContent();
    eval(UrlFetchApp.fetch('https://cdn.rawgit.com/nodeca/pako/master/dist/pako.js').getContentText());
    const binData = [];
    for (let i = 0; i < charData.length; i++) {
      binData.push(charData[i] < 0 ? charData[i] + 256 : charData[i]);
    }

    let data = pako.inflate(binData);
    let decoded_chars = '';
    for (let i = 0; i < data.length; i++) {
      decoded_chars += String.fromCharCode(data[i]);
    }
    return parentFolder.createFile(fileName, decoded_chars);
  } 
}

function callLambda(options) {
  Logger.log('calling lambda');
  apiUrl = LAMBDA_URL;
  Logger.log(options)
  const lambdaResponse = UrlFetchApp.fetch(apiUrl, options).getContentText();
  Logger.log(['lambda response: ', lambdaResponse])
  return lambdaResponse;
}
