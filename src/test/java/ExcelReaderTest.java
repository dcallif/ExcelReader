import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class ExcelReaderTest {
    private String XLS_FILE = "src/main/resources/Test.xls";
    private String XLSX_FILE = "src/main/resources/Test.xlsx";

    @Test
    @DisplayName("Test for XLS")
    void testXls() {
        assertAll("XLS Cell Tests",
                () -> {
                    ExcelReader test = new ExcelReader();
                    Workbook wb = test.getWorkbook(XLS_FILE);
                    assertNotNull(wb);

                    Sheet sheet = wb.getSheetAt(0);
                    assertNotNull(sheet);

                    assertAll("String Cell Values",
                            () -> assertTrue(test.getStrExcel(0, 3, sheet).equals("123456")),
                            () -> assertTrue(test.getStrExcel(1, 2, sheet).equals("Luna")),
                            () -> assertTrue(test.getStrExcel(1, 3, sheet).equals("true"))
                    );
                },
                () -> {
                    ExcelReader test = new ExcelReader();
                    Workbook wb = test.getWorkbook(XLS_FILE);
                    assertNotNull(wb);

                    Sheet sheet = wb.getSheetAt(0);
                    assertNotNull(sheet);

                    assertAll("Int Cell Values",
                            () -> assertTrue(test.getIntExcel(0, 3, sheet).equals(123456)),
                            () -> assertNull(test.getIntExcel(1, 2, sheet))
                    );
                }
        );
    }

    @Test
    @DisplayName("Test for XLSX")
    void testXlsx() {
        assertAll("XLS Cell Tests",
                () -> {
                    ExcelReader test = new ExcelReader();
                    Workbook wb = test.getWorkbook(XLSX_FILE);
                    assertNotNull(wb);

                    Sheet sheet = wb.getSheetAt(0);
                    assertNotNull(sheet);

                    assertAll("String Cell Values",
                            () -> assertTrue(test.getStrExcel(0, 3, sheet).equals("123456")),
                            () -> assertTrue(test.getStrExcel(1, 2, sheet).equals("Luna")),
                            () -> assertTrue(test.getStrExcel(1, 3, sheet).equals("true"))
                    );
                },
                () -> {
                    ExcelReader test = new ExcelReader();
                    Workbook wb = test.getWorkbook(XLSX_FILE);
                    assertNotNull(wb);

                    Sheet sheet = wb.getSheetAt(0);
                    assertNotNull(sheet);

                    assertAll("Int Cell Values",
                            () -> assertTrue(test.getIntExcel(0, 3, sheet).equals(123456)),
                            () -> assertNull(test.getIntExcel(1, 2, sheet))
                    );
                }
        );
    }
}