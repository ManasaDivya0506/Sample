package Smaple.Smaple;
import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
public class ExcelReaderObjConv {

    public static void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream("C:\\Playwright\\Input Excel.xlsx");

            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(0);

            Map<Integer, Map<Integer, Object>> dataMap = new HashMap<>();

            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                Map<Integer, Object> rowData = new HashMap<>();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    //switch (cell.getCellType()) {
                      //  case Cell.CELL_TYPE_NUMERIC:
                        //    rowData.put(cell.getColumnIndex(), cell.getNumericCellValue());
                         //   break;
                            switch (cell.getCellType()) {
	                        case STRING:
                            rowData.put(cell.getColumnIndex(), cell.getStringCellValue());
                            break;
	                        case NUMERIC:
	                        	rowData.put(cell.getColumnIndex(), cell.getNumericCellValue());
	                            break;
                    }
                }
                dataMap.put(row.getRowNum(), rowData);
            }

            Object[][] dataArray = new Object[dataMap.size()][dataMap.get(0).size()];

            for (int row = 0; row < dataArray.length; row++) {
                for (int col = 0; col < dataArray[row].length; col++) {
                    if (dataMap.get(row).get(col) instanceof Double) {
                        dataArray[row][col] = (Number) dataMap.get(row).get(col);
                    } else {
                        dataArray[row][col] = (String) dataMap.get(row).get(col);
                    }
                }
            }

            workbook.close();
            file.close();

            // Handle the 2D object array as needed
            for (Object[] row : dataArray) {
                for (Object cell : row) {
                    System.out.print(cell + "\t");
                }
                System.out.println();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
