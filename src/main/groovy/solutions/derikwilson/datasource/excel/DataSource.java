package solutions.derikwilson.datasource.excel;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class DataSource {

    public List<Map<String, String>> setPropertiesFromExcel(List<String> propertyNames, String filePath, String sheetName, int startRow, int startCol) {
        List<Map<String, String>> listOfPropertyMaps = new ArrayList<>();
        char[][][] excelData = loadDataFromExcel(filePath, sheetName, startRow, startCol);

        // Iterate over each row of data
        for (char[][] row : excelData) {
            Map<String, String> propertyMap = new LinkedHashMap<>(); // LinkedHashMap to preserve the order
            for (int j = 0; j < row.length; j++) {
                if (j < propertyNames.size()) {
                    String propertyName = propertyNames.get(j);
                    String propertyValue = new String(row[j]).trim();
                    propertyMap.put(propertyName, propertyValue);
                    // Clear the char[] from memory
                    Arrays.fill(row[j], '0');
                }
            }
            // Add the populated map to the list
            listOfPropertyMaps.add(propertyMap);
        }

        // Clear the entire excelData array from memory
        for (char[][] rowData : excelData) {
            for (char[] cellData : rowData) {
                Arrays.fill(cellData, '0');
            }
        }

        return listOfPropertyMaps;
    }

    public void loadDataSource(String filePath, String sheetName, int startRow, int startCol) {
        try (FileInputStream file = new FileInputStream(new File(filePath))) {
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheet(sheetName);

            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                if (row.getRowNum() < startRow) continue; // Skip rows before startRow

                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if (cell.getColumnIndex() < startCol) continue; // Skip columns before startCol

                    // Read and process cell data
                    // For example, you can get cell value as a string
                    char[] cellValue = getCellValueAsCharArray(cell);
                    // Process or store the cell value as needed
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private char[] getCellValueAsCharArray(Cell cell) {
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().toCharArray();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue()).toCharArray();
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue()).toCharArray();
            default: return "".toCharArray();
        }
    }

    public char[][][] loadDataFromExcel(String filePath, String sheetName, int startRow, int startCol) {
        List<List<char[]>> data = new ArrayList<>();
        try (FileInputStream file = new FileInputStream(new File(filePath))) {
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheet(sheetName);

            for (Row row : sheet) {
                if (row.getRowNum() < startRow) continue; // Skip rows before startRow

                List<char[]> rowData = new ArrayList<>();
                for (int cn = startCol; cn < row.getLastCellNum(); cn++) {
                    Cell cell = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    rowData.add(getCellValueAsCharArray(cell));
                }
                data.add(rowData);
            }
        } catch (Exception e) {
            e.printStackTrace();
            // Handle exceptions or throw them as needed
        }

        return convertListToArray(data);
    }

    private char[][][] convertListToArray(List<List<char[]>> listData) {
        char[][][] arrayData = new char[listData.size()][][];
        for (int i = 0; i < listData.size(); i++) {
            List<char[]> row = listData.get(i);
            arrayData[i] = row.toArray(new char[row.size()][]);
        }
        return arrayData;
    }

    public List<String> getHeaderRow(String filePath, String sheetName, int headerRowNum) {
        List<String> headers = new LinkedList<>();
        try (FileInputStream file = new FileInputStream(new File(filePath))) {
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheet(sheetName);
            Row row = sheet.getRow(headerRowNum);

            for (Cell cell : row) {
                switch (cell.getCellType()) {
                    case STRING:
                        headers.add(cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        headers.add(String.valueOf(cell.getNumericCellValue()));
                        break;
                    default:
                        headers.add("Unsupported Cell Type");
                        break;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return headers;
    }


}
