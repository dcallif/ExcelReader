import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.*;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class ExcelReaderTest {
    private String XLS_FILE = "src/main/resources/Test.xls";
    private String XLSX_FILE = "src/main/resources/Test.xlsx";
    private Workbook WB = null;
    private ExcelReader TEST = null;

    @BeforeEach
    void init(TestInfo testInfo) throws IOException {
        TEST = new ExcelReader();
        String displayName = testInfo.getDisplayName();
        if (displayName.equals("Test for XLS")) {
            WB = TEST.getWorkbook(XLS_FILE);
            assertNotNull(WB);
        } else if (displayName.equals("Test for XLSX")) {
            WB = TEST.getWorkbook(XLSX_FILE);
            assertNotNull(WB);
        } else fail();
    }

    @Test
    @DisplayName("Test for XLS")
    void testXls() {
        Sheet sheet = WB.getSheetAt(0);
        assertNotNull(sheet);

        assertAll("String Cell Values",
                () -> assertTrue(TEST.getStrExcel(0, 3, sheet).equals("123456")),
                () -> assertTrue(TEST.getStrExcel(1, 2, sheet).equals("Luna")),
                () -> assertTrue(TEST.getStrExcel(1, 3, sheet).equals("true")),
                () -> assertNull(TEST.getStrExcel(1, 200, sheet))
        );
        assertAll("Int Cell Values",
                () -> assertTrue(TEST.getIntExcel(0, 3, sheet).equals(123456)),
                () -> assertNull(TEST.getIntExcel(1, 2, sheet)),
                () -> assertNull(TEST.getIntExcel(1, 200, sheet))
        );
        assertAll("Boolean Cell Values",
                () -> assertTrue(TEST.getBoolExcel(1, 3, sheet)),
                () -> assertFalse(TEST.getBoolExcel(1, 2, sheet)),
                () -> assertFalse(TEST.getBoolExcel(1, 4, sheet)),
                () -> assertNull(TEST.getBoolExcel(1, 200, sheet))
        );
    }

    @Test
    @DisplayName("Test for XLSX")
    void testXlsx() {
        Sheet sheet = WB.getSheetAt(0);
        assertNotNull(sheet);

        assertAll("String Cell Values",
                () -> assertTrue(TEST.getStrExcel(0, 3, sheet).equals("123456")),
                () -> assertTrue(TEST.getStrExcel(1, 2, sheet).equals("Luna")),
                () -> assertTrue(TEST.getStrExcel(1, 3, sheet).equals("true")),
                () -> assertNull(TEST.getStrExcel(1, 200, sheet))
        );
        assertAll("Int Cell Values",
                () -> assertTrue(TEST.getIntExcel(0, 3, sheet).equals(123456)),
                () -> assertNull(TEST.getIntExcel(1, 2, sheet)),
                () -> assertNull(TEST.getIntExcel(1, 200, sheet))
        );
        assertAll("Boolean Cell Values",
                () -> assertTrue(TEST.getBoolExcel(1, 3, sheet)),
                () -> assertFalse(TEST.getBoolExcel(1, 2, sheet)),
                () -> assertFalse(TEST.getBoolExcel(1, 4, sheet)),
                () -> assertNull(TEST.getBoolExcel(1, 200, sheet))
        );
    }

    @AfterEach
    void tearDown() {
        WB = null;
        TEST = null;
    }
}