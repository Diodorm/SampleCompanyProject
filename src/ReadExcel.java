import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

/**
 * Created by darora on 4/18/14.
 */
public class ReadExcel {
    //Create Workbook instance from excel sheet
    public static void main(String[] args) {
        try {
            //Get the Excel File
            File file = new File(System.getProperty("user.dir") + "\\InData");
            System.out.print(System.getProperty("user.dir") + "\\InData");
            File[] files = file.listFiles();
            System.out.println();
            for(File f: files){
                System.out.println(f.getName());
                XSSFWorkbook workbook = new XSSFWorkbook();
            }
            XSSFWorkbook workbook = new XSSFWorkbook();
            //Get the Desired sheet
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Increment over rows
            for (Row row : sheet) {
                //Iterate and get the cells from the row
                Iterator cellIterator = row.cellIterator();
                // Loop till you read all the data
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "t");
                            break;
                    }
                }
                System.out.println("");
            }

            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}