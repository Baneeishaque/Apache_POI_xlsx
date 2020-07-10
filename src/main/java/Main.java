import org.apache.commons.io.FilenameUtils;
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
        main.processExcelFile("padou.xlsx");
        main.processExcelFile("padou2.xlsx");
        main.processExcelFile("yells_july.xlsx");
    }

    private void processExcelFile(String sourceFileName) {

        processExcelFileForPricesAndNumbers(sourceFileName);
        processExcelFileForNumbers(sourceFileName);
    }

    private void processExcelFileForNumbers(String sourceFileName) {

        processExcelFileForPricesAndNumbersOrForNumbers(sourceFileName, "Numbers");
    }

    private void processExcelFileForPricesAndNumbers(String sourceFileName) {

        processExcelFileForPricesAndNumbersOrForNumbers(sourceFileName, "PricesAndNumbers");
    }

    private void processExcelFileForPricesAndNumbersOrForNumbers(String sourceFileName, String processMode) {

        int i = 0;
        HashMap<Integer, CellStyle> styleMap = new HashMap<>();

        try {

            XSSFWorkbook workbook = new XSSFWorkbook();

            XSSFSheet sheetNew = null;
            if (processMode.equals("PricesAndNumbers")) {

                sheetNew = workbook.createSheet("Prices & Numbers");

            } else if (processMode.equals("Numbers")) {

                sheetNew = workbook.createSheet("Numbers");
            }
            int rowCount = 0;

            File sourceFile = new File(getClass().getResource(sourceFileName).getFile());   //creating a new file instance
            FileInputStream fis = new FileInputStream(sourceFile);   //obtaining bytes from the file
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

                boolean validationResult = false;
                if (processMode.equals("PricesAndNumbers")) {

                    validationResult = (i > 4) && ((i % 3) != 2);

                } else if (processMode.equals("Numbers")) {

                    int reminder = i % 3;
                    validationResult = (i > 4) && (reminder != 2) && (reminder != 0);
                }

                if (validationResult) {

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
                        copyCellUsingHashMap(cell, cellNew, styleMap);
                    }
                }
                System.out.println();
                if (processMode.equals("PricesAndNumbers")) {

                    try (FileOutputStream outputStream = new FileOutputStream(FilenameUtils.removeExtension(sourceFileName) + " Prices & Numbers.xlsx")) {

                        workbook.write(outputStream);
                    }
                } else if (processMode.equals("Numbers")) {

                    try (FileOutputStream outputStream = new FileOutputStream(FilenameUtils.removeExtension(sourceFileName) + " Numbers.xlsx")) {

                        workbook.write(outputStream);
                    }
                }
            }
        } catch (Exception e) {

            e.printStackTrace();
        }
    }

    public void copyCellUsingHashMap(Cell oldCell, Cell newCell, HashMap<Integer, CellStyle> styleMap) {

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
