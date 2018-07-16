import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {
    private static final String XLS_FILE = "src/main/resources/Test.xls";
    private static final String XLSX_FILE = "src/main/resources/Test.xlsx";

    /**
     * @param spreadsheetFile
     * @return
     * @throws IOException
     */
    private static Workbook getWorkbook(String spreadsheetFile) throws IOException {
        FileInputStream fis = new FileInputStream(spreadsheetFile);
        if (spreadsheetFile.endsWith(".xls")) {
            return new HSSFWorkbook(fis);
        } else if (spreadsheetFile.endsWith(".xlsx")) {
            return new XSSFWorkbook(fis);
        }
        else return null;
    }

    /**
     * @param rowNum
     * @param colmnNum
     * @param sheet
     * @return
     */
    private static String getStrExcel(int rowNum, int colmnNum, Sheet sheet) {
        Row row = sheet.getRow(rowNum);
        Cell cell = row.getCell(colmnNum);
        if (cell == null) return null;
        if (cell.getCellType() == Cell.CELL_TYPE_STRING) return cell.getStringCellValue();
        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) return String.valueOf(( (int)cell.getNumericCellValue()));
        if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) return String.valueOf(cell.getBooleanCellValue());
        else return null;
    }

    public static void main(String[] args) {
        try {
            Workbook wb = getWorkbook(XLS_FILE);
            if (wb == null) return;
            Sheet sheet = wb.getSheetAt(0);
            System.out.println(getStrExcel(0, 3, sheet));
        } catch (IOException e) {
            System.out.println("Could not read input");
        }
    }
}