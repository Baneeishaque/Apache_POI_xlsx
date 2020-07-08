import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Iterator;

public class Main {

    public static void main(String[] args) {

        Main main = new Main();
        main.processExcelFile();
    }

    private void processExcelFile() {

        int i = 0;
        try {

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheetNew = workbook.createSheet("Prices & Numbers");
            int rowCount = 0;

            File file = new File(getClass().getResource("padou.xlsx").getFile());   //creating a new file instance
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
            //creating Workbook instance that refers to .xlsx file
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object

            //iterating over excel file
            for (Row row : sheet) {

                System.out.print(++i + " ");

//                if (i == 10) {
//
//                    break;
//                }

                if ((i > 4) && ((i % 3) != 2)) {

                    Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column

                    Row rowNew = sheetNew.createRow(rowCount++);
                    int columnCount = 0;

                    while (cellIterator.hasNext()) {

                        Cell cellNew = rowNew.createCell(columnCount++);

                        Cell cell = cellIterator.next();

                        switch (cell.getCellType()) {

                            case STRING:    //field that represents string cell type
                                System.out.print(cell.getStringCellValue() + "\t\t\t");
                                cellNew.setCellValue(cell.getStringCellValue());
                                break;

                            case _NONE:
                            case ERROR:
                            case BOOLEAN:
                            case BLANK:
                            case FORMULA:
                                break;

                            case NUMERIC:    //field that represents number cell type
                                System.out.print(cell.getNumericCellValue() + "\t\t\t");
                                cellNew.setCellValue(cell.getNumericCellValue());
                                break;

                            default:
                        }
                        copyCellUsingHashMap(cell, cellNew);
                    }
                }
                System.out.println();
                try (FileOutputStream outputStream = new FileOutputStream("Prices & Numbers.xlsx")) {

                    workbook.write(outputStream);
                }
            }
        } catch (Exception e) {

            e.printStackTrace();
        }
    }

    HashMap<Integer, CellStyle> styleMap = new HashMap<>();

    public void copyCellUsingHashMap(Cell oldCell, Cell newCell) {

        int styleHashCode = oldCell.getCellStyle().hashCode();

        CellStyle newCellStyle = styleMap.get(styleHashCode);

        if (newCellStyle == null) {

            newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            styleMap.put(styleHashCode, newCellStyle);
        }
        newCell.setCellStyle(newCellStyle);
    }
}
