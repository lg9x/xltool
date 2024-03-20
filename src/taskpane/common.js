export function getColumnIndexByHeaderName(headerArr, columnName) {
  for (let j = 0; j < headerArr.length; j++) {
    if (headerArr[j] === columnName) {
      return j;
    }
  }
}
