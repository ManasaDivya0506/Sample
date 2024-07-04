package Smaple.Smaple;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelReader {

    public static void main(String[] args) {
        File file = new File("C:\\Playwright\\Input Excel.xlsx");
        HashMap<String, HashMap<String, Object>> dataMap = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(file);
                Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);

            for (int i = 1; i <= sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    HashMap<String, Object> rowData = new HashMap<>();

                    for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                        Cell cell = row.getCell(j);
                        String key = sheet.getRow(0).getCell(j).getStringCellValue();
                        Object value = getCellValue(cell);

                        rowData.put(key, value);
                    }

                    dataMap.put("Row" + i, rowData);
                }
            }

            System.out.println(dataMap);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static Object getCellValue(Cell cell) {
        DataFormatter formatter = new DataFormatter();

        switch (cell.getCellType()) {
            case NUMERIC:
                return cell.getNumericCellValue();
            case STRING:
                return cell.getStringCellValue();
            default:
                return null;
        }
    }
}