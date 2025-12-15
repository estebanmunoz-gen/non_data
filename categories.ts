// @ts-ignore
function main(workbook: ExcelScript.Workbook) {
  const date = new Date();
  const nameInventorySheet = `Inventory Report ${(date.getMonth() + 1) > 9 ? date.getMonth() + 1 : '0' + date.getMonth()}${date.getDate() > 9 ? date.getDate() : '0' + date.getDate()}${date.getFullYear().toString().substring(2)}`;
  //workbook.getWorksheet('Inventory Report').setName(nameInventorySheet);
  const Inventory_Sheet = workbook.getWorksheet(nameInventorySheet);

  const lastRow = Inventory_Sheet.getUsedRange().getLastRow().getRowIndex() + 1;
  

  // - - - - - - 
  const uriCategoryNames = 'https://raw.githubusercontent.com/estebanmunoz-gen/non_data/refs/heads/main/Helper.tsv';

  const categoryMap = new Map<number, string>();

  const fetching = async () => {
  }
  /*const response = await fetch(uriCategoryNames, {
    method: 'GET'
  });*/

  

  fetch(uriCategoryNames, {
    method: 'GET'
  }).then( async (response) => {
    const categories = await response.text();
    const rows = categories.split('\n'); 

    rows.forEach((row, index) => {
      if (index == 0) {
        return;
      }

      const catName = row.split('\t')[0].trim();
      const SUPC = parseInt(row.split('\t')[1])

      categoryMap.set(SUPC, catName);
    });

    const supcs = Inventory_Sheet.getRange(`E8:E${lastRow}`).getValues();

    supcs.forEach((row, index) => {
      const s = parseInt(row[0].toString());
      supcs[index][0] = categoryMap.get(s);
    });

    console.log(supcs);

    Inventory_Sheet.getRange(`D8:D${lastRow}`).setValues(supcs);
  });

  //fetching();
}
