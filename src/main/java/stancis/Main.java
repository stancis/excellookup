package stancis;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class Main {

  public static void main(String[] args) throws IOException {
    findData("Delete_MfgDrug_404Request", "Data1");
//    findData("Get_Drug_404Request", "Data1");
  }

  private static void findData(String rowName, String columnName) throws IOException {
    String cellValue = getCellValue(rowName, columnName, "testmatrix.xlsx");
    System.out.println(cellValue);
  }

  private static String getCellValue(String rowName, String columnName, String filename) throws IOException {
    Workbook workbook = WorkbookFactory.create(new FileInputStream(filename));
    Sheet sheet = workbook.getSheetAt(0);
    Row firstRow = sheet.getRow(0);
    int rowNameColIdx = findColumnIdx("Automation TC", firstRow);
    int colNameColIdx = findColumnIdx(columnName, firstRow);

    Iterator<Row> rowIterator = sheet.iterator();
    rowIterator.next();
    while (rowIterator.hasNext()) {
      Row row = rowIterator.next();
      if (rowName.equals(row.getCell(rowNameColIdx).getStringCellValue())) {
        return row.getCell(colNameColIdx).getStringCellValue();
      }
    }
    return null;
  }

  private static int findColumnIdx(String text, Row row) {
    for (Cell cell : row) {
      if (text.equals(cell.getStringCellValue())) {
        return cell.getColumnIndex();
      }
    }
    return -1;
  }
}
