
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PrintCell {

  static double get_final_val(double start, double end){
    return end-start;
  }
  public static void main(String[] args) {
    try {
      File file = new File("E:\\OTHERS\\BTech\\Intern-1\\IMAP\\Java Mail\\src\\student.xlsx"); // creating a new file instance
      FileInputStream fis = new FileInputStream(file); // obtaining bytes from the file
      // creating Workbook instance that refers to .xlsx file
      XSSFWorkbook wb = new XSSFWorkbook(fis);
      XSSFSheet sheet = wb.getSheetAt(0); // creating a Sheet object to retrieve object
      Iterator<Row> itr = sheet.iterator(); // iterating over excel file
      while (itr.hasNext()) {
        Row row = itr.next();
        int rowIndex=0;
        Iterator<Cell> cellIterator = row.cellIterator(); // iterating over each column
        double start_date=0, end_date=0;
        while (cellIterator.hasNext()) {
          Cell cell = cellIterator.next();
          // Switch case variable to
          // get the columnIndex
          int columnIndex = cell.getColumnIndex();
          rowIndex=cell.getRowIndex();
          // Depends upon the cell contents we need to
          // typecast
          switch(columnIndex){
            case 2:
              if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
                start_date=cell.getNumericCellValue();
              }
            case 3:
              if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
                end_date=cell.getNumericCellValue();
              }  
          }
        }
        
        try{
          FileOutputStream fis1 = new FileOutputStream(file);
          Cell cell   = row.createCell(4);
          if(rowIndex!=0){
            cell.setCellValue(get_final_val(start_date, end_date));
          }
          else{
            cell.setCellValue("diff time");
          }
          wb.write(fis1);
        }catch(Exception e) {
            System.out.println(e.getMessage());  
        }
        
        Iterator<Cell> cellIterator1 = row.cellIterator(); // iterating over each column
        while (cellIterator1.hasNext()) {
          Cell cell = cellIterator1.next();
          switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING: // field that represents string cell type
              System.out.print(cell.getStringCellValue() + "\t\t");
              break;
            case Cell.CELL_TYPE_NUMERIC: // field that represents number cell type
              System.out.print(cell.getNumericCellValue() + "\t\t");
              break;
          }
        }
        System.out.println("");
      }
      wb.close();
    } catch (Exception e) {
      e.printStackTrace();
    }
  }
}