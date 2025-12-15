// @ts-ignore
function main(workbook: ExcelScript.Workbook) {
  const date = new Date();
  const nameInventorySheet = `Inventory Report ${(date.getMonth() + 1) > 9 ? date.getMonth() + 1 : '0' + date.getMonth()}${date.getDate() > 9 ? date.getDate() : '0' + date.getDate()}${date.getFullYear().toString().substring(2)}`;
  workbook.getWorksheet('Inventory Report').setName(nameInventorySheet);
  const Inventory_Sheet = workbook.getWorksheet(nameInventorySheet);

  const lastRow = Inventory_Sheet.getUsedRange().getLastRow().getRowIndex() + 1;

  // Ocultar F:I, L:U, W:Y, AA:AK, AM:AP
  Inventory_Sheet.getRange('F:I').getEntireColumn().setColumnHidden(true);
  Inventory_Sheet.getRange('L:U').getEntireColumn().setColumnHidden(true);
  Inventory_Sheet.getRange('W:Y').getEntireColumn().setColumnHidden(true);
  Inventory_Sheet.getRange('AA:AK').getEntireColumn().setColumnHidden(true);
  Inventory_Sheet.getRange('AM:AP').getEntireColumn().setColumnHidden(true);

  // Cortar de '4 Week Cust Avg' (AL) a K
  const _4Week = Inventory_Sheet.getRange(`AL1:AL${lastRow}`).getValues();
  Inventory_Sheet.getRange('K1').getEntireColumn().insert(ExcelScript.InsertShiftDirection.right);
  Inventory_Sheet.getRange(`K1:K${lastRow}`).setValues(_4Week);
  Inventory_Sheet.getRange('AM1').getEntireColumn().delete(ExcelScript.DeleteShiftDirection.left);

  // Cortar de AA a AV
  const datosAA = Inventory_Sheet.getRange(`AA1:AA${lastRow}`).getValues();
  Inventory_Sheet.getRange(`AV1:AV${lastRow}`).setValues(datosAA);
  Inventory_Sheet.getRange('AA1').getEntireColumn().delete(ExcelScript.DeleteShiftDirection.left);

  // Insertar columna en L
  Inventory_Sheet.getRange('L1').getEntireColumn().insert(ExcelScript.InsertShiftDirection.right);

  // Agregar L = J/K
  Inventory_Sheet.getRange(`L7`).setValue('QOH vs 4 wks usage');
  Inventory_Sheet.getRange(`L8:L${lastRow}`).setFormula("=J8/K8");

  // Cortar de "Total Confirmed" a M
  const datosTotalConfirmed = Inventory_Sheet.getRange(`AU1:AU${lastRow}`).getValues();
  Inventory_Sheet.getRange('M1').getEntireColumn().insert(ExcelScript.InsertShiftDirection.right);
  Inventory_Sheet.getRange(`M1:M${lastRow}`).setValues(datosTotalConfirmed);
  Inventory_Sheet.getRange('AV1').getEntireColumn().delete(ExcelScript.DeleteShiftDirection.left);

  // Cortar de AU a N
  const datosAU = Inventory_Sheet.getRange(`AU1:AU${lastRow}`).getValues();
  Inventory_Sheet.getRange('N1').getEntireColumn().insert(ExcelScript.InsertShiftDirection.right);
  Inventory_Sheet.getRange(`N1:N${lastRow}`).setValues(datosAU);
  Inventory_Sheet.getRange('AV1').getEntireColumn().delete(ExcelScript.DeleteShiftDirection.left);

  // Fondo verde de J:N
  Inventory_Sheet.getRange('J:N').getEntireColumn().getFormat().getFill().setColor('#D8E4BC');

  // Columna nueva en D
  Inventory_Sheet.getRange('D1').getEntireColumn().insert(ExcelScript.InsertShiftDirection.right);
  Inventory_Sheet.getRange('D7').setValue('Category name');

  //Fetch category names
  /*
  if (workbook.getWorksheet('Helper_1')) {
    workbook.getWorksheet('Helper_1').delete();
  }

  workbook.addWorksheet('Helper_1');

  const Helper_Sheet = workbook.getWorksheet('Helper_1');
  Helper_Sheet.getRange('A1:C1').setValues([['Category Name', 'SUPC', 'Total Categorías']]);
  Helper_Sheet.getRange('A1:C1').getFormat().getFont().setBold(true);
  Helper_Sheet.getRange('C2').setFormula('=\'[Helper.xlsx]Categories\'!C2');

  const totalCategorias = Helper_Sheet.getRange('C2').getValue() as number;
  Helper_Sheet.getRange(`A2:B${totalCategorias + 1}`).setFormula('=\'[Helper.xlsx]Categories\'!A2');
  Helper_Sheet.getRange('A1:C1').getFormat().autofitColumns();

  const SUPCn = Inventory_Sheet.getRange(`E8:E${lastRow}`).getValues();

  for (let i in SUPCn) {
    SUPCn[i][0] = parseInt(SUPCn[i][0].toString());
  }

  Inventory_Sheet.getRange(`E8:E${lastRow}`).setValues(SUPCn);
  Inventory_Sheet.getRange(`D8:E${lastRow}`).setNumberFormat('General');

  Inventory_Sheet.getRange(`D8`).setFormula(`=XLOOKUP(E8:E${lastRow}, Helper_1!\$B\$2:\$B\$${totalCategorias + 1},Helper_1!\$A\$2:\$A\$${totalCategorias + 1})`);

  const categoriesValues = Inventory_Sheet.getRange(`D8:E${lastRow}`).getValues();
  Inventory_Sheet.getRange(`D8:E${lastRow}`).setValues(categoriesValues);

  Helper_Sheet.delete();
  */

  // ===========

  /*
  const uriCategoryNames = 'https://raw.githubusercontent.com/estebanmunoz-gen/non_data/refs/heads/main/Helper.tsv';

  const categoryMap = new Map<number, string>();

  const fetching = async () => {
    const response = await fetch(uriCategoryNames, {
      method: 'GET'
    });

    const categories = await response.text();
    const rows = categories.split('\n');

    rows.forEach((row, index) => {
      if (index > 0) {
        const catName = row.split('\t')[0].trim();
        const SUPC = parseInt(row.split('\t')[1])

        categoryMap.set(SUPC, catName);
      }
    });

    const supcs = Inventory_Sheet.getRange(`E8:E${lastRow}`).getValues();

    supcs.forEach((row, index) => {
      const s = parseInt(row[0].toString());
      supcs[index][0] = categoryMap.get(s);
    });

    console.log(supcs);

    Inventory_Sheet.getRange(`D8:D${lastRow}`).setValues(supcs);
  }
  */

  // Ordenar Sitename, Caegory y StockStatus (A → Z)
  let sortFieldSiteName: ExcelScript.SortField = {
    key: 1,
    ascending: true,
    sortOn: ExcelScript.SortOn.value
  };

  let sortFieldCategoryName: ExcelScript.SortField = {
    key: 3,
    ascending: true,
    sortOn: ExcelScript.SortOn.value
  };


  let sortFieldStockAZ: ExcelScript.SortField = {
    key: 48,
    ascending: true,
    sortOn: ExcelScript.SortOn.value
  };

  Inventory_Sheet.getRange(`A7:AW${lastRow}`).getSort().apply(
    [sortFieldCategoryName, sortFieldSiteName],
    false,
    true,
    ExcelScript.SortOrientation.rows
  );

  // Ordenar A → N → I
  const order = ['A', 'N', 'I'];
  const stockRange = Inventory_Sheet.getRange(`A8:AW${lastRow}`);
  const stockValues = stockRange.getValues();
  const sortColumn = 48;

  stockValues.sort((a, b) => {
    const valA = a[sortColumn] as string;
    const valB = b[sortColumn] as string;

    const indexA = order.indexOf(valA);
    const indexB = order.indexOf(valB);

    if (indexA < indexB) {
      return -1;
    } else if (indexA > indexB) {
      return 1;
    } else {
      return 0;
    }
  });

  stockRange.setValues(stockValues);



  // FORMATO

  // Header borders
  Inventory_Sheet.getRange('AW7').getFormat().getFont().setBold(true);
  Inventory_Sheet.getRange('AW7').getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setColor('#000000');
  Inventory_Sheet.getRange('AW7').getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setColor('#000000');

  // Column width
  Inventory_Sheet.getRange('L7:N7').getEntireColumn().getFormat().setColumnWidth(50);
  Inventory_Sheet.getRange('O7').getEntireColumn().getFormat().setColumnWidth(230);

  Inventory_Sheet.getRange(`L7:M${lastRow}`).setNumberFormat('0.00');

  // Color Fill en A → F
  let evenRule = Inventory_Sheet.getRange(`A8:F${lastRow}`).addConditionalFormat(
    ExcelScript.ConditionalFormatType.custom
  );
  evenRule.getCustom().getRule().setFormula(`=MOD(ROW(),2)=0`);
  evenRule.getCustom().getFormat().getFill().setColor("#F4EDCA");

  let oddRule = Inventory_Sheet.getRange(`A8:F${lastRow}`).addConditionalFormat(
    ExcelScript.ConditionalFormatType.custom
  );
  oddRule.getCustom().getRule().setFormula(`=MOD(ROW(),2)=1`);
  oddRule.getCustom().getFormat().getFill().setColor("#FFFFFF");
  
  // Color Fill en O → AW
  evenRule = Inventory_Sheet.getRange(`P8:AW${lastRow}`).addConditionalFormat(
    ExcelScript.ConditionalFormatType.custom
  );
  evenRule.getCustom().getRule().setFormula(`=MOD(ROW(),2)=0`);
  evenRule.getCustom().getFormat().getFill().setColor("#F4EDCA");

  oddRule = Inventory_Sheet.getRange(`P8:AFW${lastRow}`).addConditionalFormat(
    ExcelScript.ConditionalFormatType.custom
  );
  oddRule.getCustom().getRule().setFormula(`=MOD(ROW(),2)=1`);
  oddRule.getCustom().getFormat().getFill().setColor("#FFFFFF");
}