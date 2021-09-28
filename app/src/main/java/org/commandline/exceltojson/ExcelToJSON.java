package org.commandline.exceltojson;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


/*
   Taken from an example of using the Apache POI api at: https://www.dev2qa.com/convert-excel-to-json-in-java-example/
 */

public class ExcelToJSON {
    private File excelFile;

    public ExcelToJSON(File testFile) {
        excelFile = testFile;
    }

    //Note: this assumes all data is on Sheet zero
    public String toJson() {
        StringBuilder sb = new StringBuilder();
        //With resources try-catch should auto close
        //TODO: Might be fun to throw garbage at this just to confirm
        try (FileInputStream fis = new FileInputStream(excelFile)) {
            XSSFWorkbook workBook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workBook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            List<List<String>> sheetDataTable = loadSheetDataTable(sheet);
            sb.append(exportJSONStringFromSheetData(sheetDataTable));

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return sb.toString();
    }

    /* Return sheet data in a two dimensional list.
     * Each element in the outer list is represent a row,
     * each element in the inner list represent a column.
     * The first row is the column name row.*/
    private static List<List<String>> loadSheetDataTable(XSSFSheet sheet) {
        List<List<String>> ret = new ArrayList<List<String>>();
        // Get the first and last sheet row number.
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum > 0) {
            // Loop in sheet rows.
            for (int i = firstRowNum; i < lastRowNum + 1; i++) {
                // Get current row object.
                Row row = sheet.getRow(i);
                // Get first and last cell number.
                int firstCellNum = row.getFirstCellNum();
                int lastCellNum = row.getLastCellNum();
                // Create a String list to save column data in a row.
                List<String> rowDataList = new ArrayList<String>();
                // Loop in the row cells.
                for (int j = firstCellNum; j < lastCellNum; j++) {
                    // Get current cell.
                    Cell cell = row.getCell(j);
                    // Get cell type.
                    CellType cellType = cell.getCellType();
                    if (cellType == CellType.NUMERIC) {
                        double numberValue = cell.getNumericCellValue();
                        // BigDecimal is used to avoid double value is counted use Scientific counting method.
                        // For example the original double variable value is 12345678, but jdk translated the value to 1.2345678E7.
                        String stringCellValue = BigDecimal.valueOf(numberValue).toPlainString();
                        rowDataList.add(stringCellValue);
                    } else if (cellType == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        rowDataList.add(cellValue);
                    } else if (cellType == CellType.BOOLEAN) {
                        boolean numberValue = cell.getBooleanCellValue();
                        String stringCellValue = String.valueOf(numberValue);
                        rowDataList.add(stringCellValue);
                    } else if (cellType == CellType.BLANK) {
                        rowDataList.add("");
                    }
                }
                // Add current row data list in the return list.
                ret.add(rowDataList);
            }
        }
        return ret;
    }

    /* Return a JSON string from the string list. */
    private static String exportJSONStringFromSheetData(List<List<String>> dataTable)
    {
        String ret = "";
        if(dataTable != null)
        {
            int rowCount = dataTable.size();
            if(rowCount > 1)
            {
                // Create a JSONObject to store table data.
                JSONObject tableJsonObject = new JSONObject();
                // The first row is the header row, store each column name.
                List<String> headerRow = dataTable.get(0);
                int columnCount = headerRow.size();
                // Loop in the row data list.
                for(int i=1; i<rowCount; i++)
                {
                    // Get current row data.
                    List<String> dataRow = dataTable.get(i);
                    // Create a JSONObject object to store row data.
                    JSONObject rowJsonObject = new JSONObject();
                    for(int j=0;j<columnCount;j++)
                    {
                        String columnName = headerRow.get(j);
                        String columnValue = dataRow.get(j);
                        rowJsonObject.put(columnName, columnValue);
                    }
                    tableJsonObject.put("Row " + i, rowJsonObject);
                }
                // Return string format data of JSONObject object.
                ret = tableJsonObject.toString();
            }
        }
        return ret;
    }
}
