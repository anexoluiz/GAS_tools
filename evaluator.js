function onChange(e) {
  const range = e.source.getActiveSheet().getActiveRange();
  const rangearray = range.getValues();
  const firstcol = range.getColumn();
  const lastCol = range.getLastColumn();
  let colrange;
  if (firstcol == lastCol) {
    colrange = [firstcol];
  } else {
    colrange = Array.from({ length: lastCol - firstcol + 1 }, (_, i) => firstcol + i);
  }
  const firstrow = range.getRow();
  if (firstrow > 1) {
    // Apply formatting to all range to minimize API calls
    const styles = SpreadsheetApp.newTextStyle().setForegroundColor('black').setFontSize(10).setBold(false).setItalic(false).setUnderline(false).setStrikethrough(false).setFontFamily('Arial').build();
    range.setTextStyle(styles);
    range.setBackground(null);
    range.setHorizontalAlignment('center');
  }
  let ranges = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
  let paintRedRangeList = [];
  for (row in rangearray) {
    if (colrange.includes(4)) {
      const index = colrange.indexOf(4);
      const value = rangearray[row][index].toString().replace(/\D/g, '');
      if (value.length % 11 != 0) {
        // Only exact eleven digits cellphones without symbols are allowed
        rangearray[row][index] = 'Inválido: Apenas celulares';
        paintRedRangeList.push(ranges[3] + (parseInt(firstrow) + parseInt(row)));
      } else if (rangearray[row][index].toString().length && value.length == 0) {
        // Remove values without numbers
        rangearray[row][index] = 'Inválido';
        paintRedRangeList.push(ranges[3] + (parseInt(firstrow) + parseInt(row)));
      } else if (value.length > 0 && value.length / 11 != 1) {
        // Validate just one number per cell
        rangearray[row][index] = 'Inválido: Apenas um celular por linha';
        paintRedRangeList.push(ranges[3] + (parseInt(firstrow) + parseInt(row)));
      }
    }
    if (colrange.includes(3)) {
      const index = colrange.indexOf(3);
      const value = rangearray[row][index].toString().replace(/\D/g, '');
      if (rangearray[row][index].toString().length > 0 && value.length !== 8) {
        // Identifier number must be 8-digit
        rangearray[row][index] = 'Inválido';
        paintRedRangeList.push(ranges[2] + (parseInt(firstrow) + parseInt(row)));
      }
    }
  }
  range.setValues(rangearray);
  if (paintRedRangeList.length > 0) {
    // Paint all wrong ranges at once to minimize API calls, and speed UP the execution
    let redRanges = range.getSheet().getRangeList(paintRedRangeList);
    redRanges.setBackground('red');
    redRanges.setFontWeight('bold');
  }
}
