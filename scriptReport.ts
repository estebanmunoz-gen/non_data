interface Row {
  SUPC: number;
  item_description: string;
  QOH: number;
  _4week: number;
  total_confirmed: number;

  openPOs: string;
  qoc: number;
  source_vendor: string;
  sale_date: number;
  recvd_qty: number;
  recvd_date: number;
  stock_status: string;
}

// @ts-ignore
function main(workbook: ExcelScript.Workbook) {
  // == Globals ==

  // Indexes
  let currentCategorieIndex = 0;
  let currentRowIndex = 7;
  let currentZone = '';

  const categories = ['BEEF', 'CHICKEN', 'PORK', 'SEAFOOD', 'SAUCE', 'BANCHAN', 'PRODUCE', 'OTHERS'];

  const worldMap = new Map<string, Map<string, Row[]>>();

  const date = new Date();
  const saturday = new Date();
  const bussinesDay = new Date();

  saturday.setDate(date.getDate() - date.getDay() + 7);
  bussinesDay.setDate(date.getDate() - date.getDay() + 7 + (date.getDay() > 1 ? 5 : 0));


  const inventoryNameSheet = `Inventory Report ${(date.getMonth() + 1) > 9 ? date.getMonth() + 1 : '0' + date.getMonth()}${date.getDate() > 9 ? date.getDate() : '0' + date.getDate()}${date.getFullYear().toString().substring(2)}`;
  const InventoryReport_Sheet = workbook.getWorksheet(inventoryNameSheet);
  const lastRowInventory = InventoryReport_Sheet.getUsedRange().getLastRow().getRowIndex() + 1;

  // Delete ALL none used sheets
  workbook.getWorksheets().forEach( (sheet) => {
    if (![inventoryNameSheet, 'Legend'].includes(sheet.getName())) {
      sheet.delete();
    }
  });

  // Functions = = = = = = = =
  const createStatic = (zone: string, numberZone: string = '') => {
    const ActualSheet = workbook.getWorksheet(zone);
    const headers = [['Category name', 'SUPC', 'Item Description', 'QOH', '4 Week Cust Avg', 'QOH vs 4 wks usage', 'Total confirmed Qty this week(Mon - Fri)', 'QOH vs 4 wks usage', 'Open POs(PO#, Rec Date, Qty on Order, Qy Conf)', 'QOO', 'Source Vendor Name', 'Site Last Sale Date', 'Site Last Recvd Qty', 'Site Last Recvd Date', 'Stock Status']];

    ActualSheet.getRange('A6:O6').setValues(headers)
    ActualSheet.getRange('A6:O6').getFormat().getFont().setBold(true);
    ActualSheet.getRange('A6:O6').getEntireRow().getFormat().setRowHeight(48);
    ActualSheet.getRange('A6:O6').getEntireRow().getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    ActualSheet.getRange('A6:O6').getEntireColumn().getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    ActualSheet.getRange('A3').setValue(`DC #${numberZone.padStart(3, '0')} - ${zone}`);
    ActualSheet.getRange('A2:A3').getFormat().getFont().setSize(18);
    ActualSheet.getRange('A2').setValue('mySysco Product Inventory Report');
    ActualSheet.getRange('A5').setValue('    - GLEN ROSE & BBQ & RED CHAMBER & WANG');
    currentRowIndex--;
    setNormalBorders(ActualSheet);
    currentRowIndex++;
  }

  const mapAll = () => {
    const allData = InventoryReport_Sheet.getRange(`A8:AW${lastRowInventory}`).getValues();

    allData.forEach((row) => {
      const zone = row[1] as string;
      const categoryName = row[3] as string;

      const obj: Row = {
        SUPC: row[4] as number,
        item_description: row[5] as string,
        QOH: row[10] as number,
        _4week: row[11] as number,
        total_confirmed: row[13] as number,

        openPOs: row[14] as string,
        qoc: row[15] as number,
        source_vendor: row[26] as string,
        sale_date: row[45] as number,
        recvd_qty: row[46] as number,
        recvd_date: row[47] as number,
        stock_status: row[48] as string
      }

      if (!worldMap.has(zone)) {
        worldMap.set(zone, new Map<string, Row[]>());
      }

      const zoneMap = worldMap.get(zone);

      if (!zoneMap.has(categoryName)) {
        zoneMap.set(categoryName, []);
      }

      const items = zoneMap.get(categoryName);
      items.push(obj);
      zoneMap.set(categoryName, items);
      worldMap.set(zone, zoneMap);
    });
  }

  const setDisBorders = (Working_Sheet: ExcelScript.Worksheet, first: number, last: number) => {
    if (first == -1) {
      return;
    }

    const range = Working_Sheet.getRange(`A${first}:O${last}`).getFormat();

    range.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setColor('black');
    range.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setColor('black');
    range.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setColor('black');
    range.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setColor('black');
    range.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.medium);
    range.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.medium);
    range.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.medium);
    range.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.medium);
  }

  const setNormalBorders = (Working_Sheet: ExcelScript.Worksheet) => {
    const range = Working_Sheet.getRange(`A${currentRowIndex}:O${currentRowIndex}`).getFormat();
    range.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setColor('black');
    range.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setColor('black');
    range.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setColor('black');
    range.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setColor('black');
    range.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setColor('black');
    range.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
    range.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
    range.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
    range.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
    range.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
  }

  const setItems = (Working_Sheet: ExcelScript.Worksheet) => {
    const zoneCatValues = worldMap.get(currentZone).get(categories[currentCategorieIndex]);
    let i = 1;
    let firstDis = -1;
    let lastDis = -1;
    zoneCatValues.forEach((value, index) => {
      const values: (number | string)[][] = [[categories[currentCategorieIndex], value.SUPC, value.item_description, value.QOH, value._4week, 'CHANGE FOR FORMULA', value.total_confirmed, 'CHANGE FOR FORMULA', value.openPOs.replace('  ', '\n'), value.qoc, value.source_vendor, value.sale_date, value.recvd_qty, value.recvd_date, value.stock_status == 'I' ? 'DISCONTINUED' : value.stock_status]];

      if (value.stock_status == 'I') {
        if (i == 1) {
          firstDis = currentRowIndex;
        }
        values[0][0] = i.toString();
        i++

        if (index == zoneCatValues.length - 1) {
          lastDis = currentRowIndex;
        }

        Working_Sheet.getRange(`O${currentRowIndex}`).getFormat().getFont().setColor('#FF0000');
      }

      Working_Sheet.getRange(`A${currentRowIndex}:O${currentRowIndex}`).setValues(values);

      Working_Sheet.getRange(`F${currentRowIndex}`).setFormula(`=D${currentRowIndex}/E${currentRowIndex}`);
      Working_Sheet.getRange(`H${currentRowIndex}`).setFormula(`=(D${currentRowIndex}+G${currentRowIndex})/E${currentRowIndex}`);

      Working_Sheet.getRange(`D${currentRowIndex}:I${currentRowIndex}`).getFormat().getFill().setColor('#D8E4BC');

      // const fv1 = Working_Sheet.getRange(`F${currentRowIndex}`).getValue().toString();
      // const fv2 = Working_Sheet.getRange(`H${currentRowIndex}`).getValue().toString();

      Working_Sheet.getRange(`F${currentRowIndex}`).setNumberFormat('0.00');
      Working_Sheet.getRange(`H${currentRowIndex}`).setNumberFormat('0.00');

      verifyDate(Working_Sheet);
      Working_Sheet.getRange(`A${currentRowIndex}:O${currentRowIndex}`).getFormat().getFont().setSize(10);

      setNormalBorders(Working_Sheet);
      currentRowIndex++;
    });
    setDisBorders(Working_Sheet, firstDis, lastDis);
  }

  const verifyDate = (Working_Sheet: ExcelScript.Worksheet) => {
    const rawCell = Working_Sheet.getRange(`I${currentRowIndex}`);
    const raw = rawCell.getValue().toString().trim();
    const today = new Date();

    if (raw == '') {
      Working_Sheet.getRange(`A${currentRowIndex}`).getEntireRow().getFormat().setRowHeight(25.5);
      return;
    }

    const dates = raw.split('\n');
    let sumatory = 0;
    const charIndex = 0;
    dates.forEach((dateR) => {
      const fv1 = Working_Sheet.getRange(`F${currentRowIndex}`).getValue().toString();
      const fv2 = Working_Sheet.getRange(`H${currentRowIndex}`).getValue().toString();
      const dateRow = new Date(dateR.split(',')[1]);

      if (dateRow.getTime() < (saturday.getTime() + 1000)) {
        sumatory += parseInt(dateR.split(',')[3]);
      }

      if ((dateRow.getTime() < (bussinesDay.getTime() + 1000)) && (parseFloat(fv1) < 3)) {
        Working_Sheet.getRange(`I${currentRowIndex}`).getFormat().getFont().setColor('red');
      }
    });

    Working_Sheet.getRange(`A${currentRowIndex}`).getEntireRow().getFormat().autofitRows();

    if (Working_Sheet.getRange(`A${currentRowIndex}`).getEntireRow().getFormat().getRowHeight() < 27) {
      Working_Sheet.getRange(`A${currentRowIndex}`).getEntireRow().getFormat().setRowHeight(27);
    }

    //console.log(Working_Sheet.getRange(`A${lastRow}:A${lastRow}`).getEntireRow().getFormat().autofitRows());
    //Working_Sheet.getRange(`A7:A${lastRow}`).getEntireRow().getFormat().autofitRows();

    Working_Sheet.getRange(`G${currentRowIndex}`).setValue(sumatory);
  }

  const setCategories = (Working_Sheet: ExcelScript.Worksheet) => {
    const categorieName = categories[currentCategorieIndex];
    Working_Sheet.getRange(`A${currentRowIndex}`).setValue(categorieName);
    Working_Sheet.getRange(`A${currentRowIndex}`).getFormat().getFont().setSize(12);
    Working_Sheet.getRange(`A${currentRowIndex}`).getFormat().getFont().setBold(true);
    Working_Sheet.getRange(`A${currentRowIndex}`).getFormat().setRowHeight(27);
    currentRowIndex++;
    setItems(Working_Sheet);

    if (currentCategorieIndex < 7) {
      currentCategorieIndex++;
      setCategories(Working_Sheet);
    } else {
      currentCategorieIndex = 0;
      currentRowIndex = 7;
    }
  }

  const setColumnsWidth = (Working_Sheet: ExcelScript.Worksheet) => {
    const lastRow = Working_Sheet.getUsedRange().getLastRow().getRowIndex() + 1
    Working_Sheet.getRange('A1').getEntireColumn().getFormat().setColumnWidth(101.6);
    Working_Sheet.getRange('B1').getEntireColumn().getFormat().setColumnWidth(51.6);
    Working_Sheet.getRange('C1').getEntireColumn().getFormat().setColumnWidth(210.6);
    Working_Sheet.getRange('D1:H1').getEntireColumn().getFormat().setColumnWidth(70);
    Working_Sheet.getRange('J1:O1').getEntireColumn().getFormat().setColumnWidth(100);
    Working_Sheet.getRange('D1:O1').getEntireColumn().getFormat().setWrapText(true);
    Working_Sheet.getRange('I1').getEntireColumn().getFormat().setColumnWidth(150);
    Working_Sheet.getRange('A1:O1').getEntireColumn().getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);

    Working_Sheet.getRange('A1').getEntireColumn().getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    Working_Sheet.getUsedRange().getFormat().getFont().setName('Arial');

    const qohRange = Working_Sheet.getRange(`F8:F${lastRow}`);
    const qohRange2 = Working_Sheet.getRange(`H8:H${lastRow}`);
    const condRange_D = Working_Sheet.getRange(`D8:D${lastRow}`);

    // RULE D.3 BLANK
    let cond_D_B = condRange_D.addConditionalFormat(
      ExcelScript.ConditionalFormatType.presetCriteria
    ).getPreset();

    let presetD_B = cond_D_B.setRule({ criterion: ExcelScript.ConditionalFormatPresetCriterion.blanks });
    cond_D_B.getFormat().getFill().setColor('#FFFFFF');

    // RULE D.1 Red in 0
    let ruleD = {
      formula1: "=0",
      operator: ExcelScript.ConditionalCellValueOperator.equalTo
    };

    let cond_D = condRange_D.addConditionalFormat(
      ExcelScript.ConditionalFormatType.cellValue
    ).getCellValue();

    cond_D.setRule(ruleD);
    cond_D.getFormat().getFont().setColor('#FF0000');
    cond_D.getFormat().getFill().setColor('#F2DCDB');



    // RULE 1.3 BLANK
    let cond_qohB = qohRange.addConditionalFormat(
      ExcelScript.ConditionalFormatType.presetCriteria
    ).getPreset();

    let presetB = cond_qohB.setRule({ criterion: ExcelScript.ConditionalFormatPresetCriterion.blanks });
    cond_qohB.getFormat().getFill().setColor('#FFFFFF');

    // RULE 1.1
    let rule: ExcelScript.ConditionalCellValueRule = {
      formula1: "=3",
      operator: ExcelScript.ConditionalCellValueOperator.lessThan
    };

    let cond_qoh = qohRange.addConditionalFormat(
      ExcelScript.ConditionalFormatType.cellValue
    ).getCellValue();

    cond_qoh.setRule(rule);
    cond_qoh.getFormat().getFill().setColor('#F2DCDB');
    cond_qoh.getFormat().getFont().setColor('#FF0000');

    // RULE 1.2 DIV/0
    let cond_qohE = qohRange.addConditionalFormat(
      ExcelScript.ConditionalFormatType.presetCriteria
    ).getPreset();

    let presetE = cond_qohE.setRule({ criterion: ExcelScript.ConditionalFormatPresetCriterion.errors });
    cond_qohE.getFormat().getFont().setColor('#D8E4BC');


    // RULE 2.4 BLANK
    let cond_qoh2B = qohRange2.addConditionalFormat(
      ExcelScript.ConditionalFormatType.presetCriteria
    ).getPreset();

    let preset2B = cond_qoh2B.setRule({ criterion: ExcelScript.ConditionalFormatPresetCriterion.blanks });
    cond_qoh2B.getFormat().getFill().setColor('#FFFFFF');

    // RULE 2.1 Red
    let rule2: ExcelScript.ConditionalCellValueRule = {
      formula1: "=3",
      operator: ExcelScript.ConditionalCellValueOperator.lessThan
    };

    let cond_qohH = qohRange2.addConditionalFormat(
      ExcelScript.ConditionalFormatType.cellValue
    ).getCellValue();

    cond_qohH.setRule(rule);
    cond_qohH.getFormat().getFill().setColor('#F2DCDB');
    cond_qohH.getFormat().getFont().setColor('#FF0000');

    // RULE 2.1 Blue
    rule2 = {
      formula1: "=3",
      operator: ExcelScript.ConditionalCellValueOperator.greaterThanOrEqual
    };

    cond_qohH = qohRange2.addConditionalFormat(
      ExcelScript.ConditionalFormatType.cellValue
    ).getCellValue();

    cond_qohH.setRule(rule2);
    cond_qohH.getFormat().getFont().setColor('#0070C0');

    // RULE 2.3 DIV/0
    let cond_qohHE = qohRange2.addConditionalFormat(
      ExcelScript.ConditionalFormatType.presetCriteria
    ).getPreset();

    let presetHE = cond_qohHE.setRule({ criterion: ExcelScript.ConditionalFormatPresetCriterion.errors });
    cond_qohHE.getFormat().getFont().setColor('#D8E4BC');

    Working_Sheet.getRange('A1:A5').getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  }


  // INIT = = = = = =
  if (workbook.getWorksheet('Helper')) {
    workbook.getWorksheet('Helper').delete();
  }

  workbook.addWorksheet('Helper');
  const Helper_Sheet = workbook.getWorksheet('Helper');

  // Get Zones
  // Buscar zonas únicas del inventario, y contarlas
  Helper_Sheet.getRange('A2').setFormula(`=UNIQUE('${inventoryNameSheet}'!B8:B${lastRowInventory})`);
  Helper_Sheet.getRange('B2').setFormula('=COUNTA(A:A)');

  // Obtener total de zonas
  const totalZones = parseInt(Helper_Sheet.getRange('B2').getValue().toString());

  Helper_Sheet.getRange(`C2:C${totalZones}`).setFormula(`=XLOOKUP(A2, '${inventoryNameSheet}'!B8:B1277, '${inventoryNameSheet}'!A8:A1277)`);

  // Mapear zonas con su código
  const zones = Helper_Sheet.getRange(`A2:C${totalZones + 1}`).getValues();

  // MAP EVERYTHING
  mapAll();

  // Create Zones Sheet
  zones.forEach((zone) => {
    if (workbook.getWorksheet(zone[0] as string)) {
      workbook.getWorksheet(zone[0] as string).delete();
    }
    workbook.addWorksheet(zone[0] as string);
    const WorkingSheet = workbook.getWorksheet(zone[0] as string);
    currentZone = zone[0] as string;

    createStatic(zone[0] as string, zone[2].toString());

    setCategories(WorkingSheet);
    setColumnsWidth(WorkingSheet);
  });



  // Move every sheet for their "code name" and position
  const sheetsOrder = [
    ["N TX", 'North Texas'],
    ["C TX", 'Central Texas'],
    ["LA", 'Los Angeles'],
    ["LV", 'Las Vegas'],
    ["HOUSTON", 'Houston'],
    ["S F", 'San Francisco'],
    ["AZ", 'Arizona'],
    ["LONG ISLAND", 'Long Island'],
    ["S FL", 'South Florida'],
    ["SEATTLE", 'Seattle'],
    ["RALEIGH", 'Raleigh'],
    ["JV", 'Jacksonville'],
    ["NEW MEXICO", 'New Mexico'],
    ["NASHVILLE", 'Nashville'],
  ];

  sheetsOrder.forEach((zone: string[], i) => {
    if (!workbook.getWorksheet(zone[1])) {
      return;
    }

    console.log(i, workbook.getWorksheet(zone[1]).setPosition(i + 1));
    workbook.getWorksheet(zone[1]).setName(zone[0]);
  });

  // Delete Helper Sheet
  workbook.getWorksheet('Helper').delete();
}