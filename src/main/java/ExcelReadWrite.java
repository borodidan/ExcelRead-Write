import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.Iterator;

/**
 * Created by rajeevkumarsingh on 18/12/17.
 */

public class ExcelReadWrite {
    public static final String SAMPLE_XLS_FILE_PATH = "./sample-xls-file.xls";
    public static final String SAMPLE_XLSX_FILE_PATH = "./sample-xlsx-file.xlsx";
    public static final String Write_In_File_Path = "./poi-generated-file.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        // 1. You can obtain a sheetIterator and iterate over it
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        System.out.println("Retrieving Sheets using Iterator");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }

        // 2. Or you can use a for-each loop
        System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // 1. You can obtain a rowIterator and columnIterator and iterate over them
        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        ArrayList<String> values = new ArrayList();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                System.out.println();
                String cellValue = dataFormatter.formatCellValue(cell);
                if(cell.getColumnIndex()== 4){
                    values.add(cellValue);
                }
            }
        }
        System.out.println(values.toString());
        writeInExcel(values);
        workbook.close();
    }

    public static void writeInExcel(ArrayList<String> values) throws IOException, InvalidFormatException {

        FileInputStream inputStream = new FileInputStream(new File(Write_In_File_Path));

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbookWrite = WorkbookFactory.create(inputStream);

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbookWrite.getNumberOfSheets() + " Sheets : ");

        Sheet sheet = workbookWrite.getSheetAt(0);
        int rowCount = 0;

        int columnCount= sheet.getRow(0).getPhysicalNumberOfCells();
        System.out.println(columnCount);

        for (String aBook : values) {
            Row row = sheet.getRow(rowCount);
            if(row == null){
             row=sheet.createRow(rowCount);
            }
            rowCount++;
            Cell cell = row.createCell(columnCount);
            cell.setCellValue(aBook);
        }
        inputStream.close();

        FileOutputStream outputStream = new FileOutputStream("poi-generated-file.xlsx");
        workbookWrite.write(outputStream);
        workbookWrite.close();
        outputStream.close();
    }

}
