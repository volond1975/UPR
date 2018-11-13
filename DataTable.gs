/**
 * Produce a dataTable object suitable for use with Charts, from
 * an array of rows (such as you'd get from Range.getValues()).
 * Assumes labels are in row 0, and the data types in row 1 are
 * representative for the table.
 * https://gist.github.com/mogsdad/8714493 
 * @param {Array} data  Array of table rows
 *
 @ @returns {DataTable} Refer to GAS documentation
 */
function DataTableFromArray( data ) {
Logger.log(data)
 var dataTable = Charts.newDataTable();
  for (var col=0; col<data[0].length; col++) {
    var label = data[0][col];
    var firstCell = data[1][col];
    if (typeof firstCell == 'string')
      dataTable.addColumn(Charts.ColumnType.STRING, label);
    else
      dataTable.addColumn(Charts.ColumnType.NUMBER, label);
  }
  for (var row = 1; row < data.length; row++) {
    dataTable.addRow(data[row]);
  }  
  return dataTable.build();
}
