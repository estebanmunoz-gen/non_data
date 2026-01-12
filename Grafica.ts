
function main(workbook: ExcelScript.Workbook, semana: string) {
    let selectedSheet = workbook.getActiveWorksheet();

    const testValues = selectedSheet.getUsedRange().getValues();
    let lastRow = 0;

    testValues.forEach( (row, index) => {
        if (row[2] == 'TOTAL') {
            lastRow = index;
        }
    });

    const ranks = selectedSheet.getRange(`A3:A${lastRow + 1}`).getValues();
    const names = selectedSheet.getRange(`C3:C${lastRow + 1}`).getValues();
    const totalQty = selectedSheet.getRange(`DP3:DP${lastRow + 1}`).getValues();
    const totalAmount = selectedSheet.getRange(`DQ3:DQ${lastRow + 1}`).getValues();
    const percentage = selectedSheet.getRange(`DR3:DR${lastRow + 1}`).getValues();
    const syscoPrice = selectedSheet.getRange(`DS3:DS${lastRow + 1}`).getValues();
    const manufacturer = selectedSheet.getRange(`DT3:DT${lastRow + 1}`).getValues();

    const data : (string | number | boolean)[][]= [
        ["Category", "Total Qty", "Total Amount", "Percentage", "Sysco Price", "Manufacturer", "Rank"]
    ];

    const colors : string[] = [];
    const usdFormatter = new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
    });

    for (let i = 0; i < ranks.length - 1; i++) {
        data.push([
          names[i][0],
          totalQty[i][0].toLocaleString('en-US'),
          usdFormatter.format(totalAmount[i][0] as number),
          (Math.round(percentage[i][0] as number * 1000) / 10) + '%',
          usdFormatter.format(syscoPrice[i][0] as number),
          usdFormatter.format(manufacturer[i][0] as number),
          ranks[i][0]
        ]);
        colors.push(selectedSheet.getRange(`B${3 + i}`).getFormat().getFill().getColor());

        if (colors[i].toUpperCase() == '#FFFFFF') {
          colors.pop();
          colors.push('#C9DAF8')
        }
    }

    workbook.addWorksheet('WK-' + semana + ' 2026');
    const sheet = workbook.getWorksheet('WK-' + semana + ' 2026');
    const range = sheet.getRange('A1:G' + (data.length));

    //console.log(colors);

    range.setValues(data);

    const chart = sheet.addChart(ExcelScript.ChartType.columnClustered, range);

    let dataTable = chart.getDataTable();
    if (dataTable) {
        dataTable.setVisible(true);
        dataTable.setShowLegendKey(true);
    }

    chart.getTitle().setText('GEN TOTAL QUANTITY (ALL STORES) WE-' + semana + ', 2026');
    chart.getSeries()[0].setGapWidth(100);
    chart.getSeries()[0].setOverlap(100);

  for (let i = 0; i < chart.getSeries()[0].getPoints().length; i++) {
    chart.getSeries()[0].getPoints()[i].getFormat().getFill().setSolidColor(colors[i]);
  }

    for (let i = 0; i < 6; i++) {
        let series = chart.getSeries()[i];
        series.setHasDataLabels(true);
        const dataLabels = series.getDataLabels();

        dataLabels.setShowValue(true);

        dataLabels.setShowLegendKey(false);

        dataLabels.getFormat().getFont().setColor("#000000");

        if (i > 0) {
            series.getFormat().getFill().clear();

            series.setHasDataLabels(false);
            series.getFormat().getLine().setWeight(0);
        }
    }

    chart.setWidth(3000);
    chart.setHeight(600);
    chart.getLegend().setVisible(false)

    const valueAxis = chart.getAxes().getValueAxis();

    valueAxis.setMaximum((Math.round(totalQty[0][0] as number / 3000) + 1) * 3000);

    const image = chart.getImage(3000,600, ExcelScript.ImageFittingMode.fitAndCenter)
    selectedSheet.addImage(image);

    sheet.delete();
}
