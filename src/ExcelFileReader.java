import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * This this is a simple project to reader the excel file from the source s
 */
public class ExcelFileReader {
    private static org.apache.poi.ss.usermodel.Sheet sheetAt;

    public static void main(String[] args) throws IOException {
        System.out.println("Enter the location : " ) ;
        String path = "./src/ReadSample.xlsx";
        File file = new File(path);
        System.out.println(file.getAbsolutePath());
        System.out.println(file.getPath());
        System.out.println(file.getName());
        //file.getPath();
        readExcelFile(file);

    }

    private static void readExcelFile(File file) throws IOException {
    
        FileInputStream fileInputStream = new FileInputStream(file);
            getApachePOI(fileInputStream);
    }

    private static void getApachePOI(FileInputStream fileInputStream) {
        XSSFWorkbook workbook;
        try {
            workbook = new XSSFWorkbook(fileInputStream);// get the input Stream and read as workbook
                 sheetAt = workbook.getSheetAt(0);// get the sheet in the workbooks
                 sheetAt.getLastRowNum();// the number of rows
                 System.out.println("Last Row Number :" + sheetAt.getLastRowNum());
                 
                 Iterator<Row> iterator = sheetAt.iterator();
                 while (iterator.hasNext()) {
                        
                     //   System.out.println("Iterator Row:" + iterator.next().Fir);
                         Row row = iterator.next();
                         
                         System.out.println("");
                        for (Cell rowIn : row) {
                            System.out.print("" + rowIn.getStringCellValue());
                            System.out.print("|");
                        }
                         

                 }

                

                        
        } catch (IOException e) {
            
            e.printStackTrace();
        } //Workbook
        
    }
    
}
