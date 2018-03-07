import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class ApachePOIExcelRead {

    private static final String FILE_NAME = "C:\\Users\\minht\\Desktop\\GDtrung.xlsx";
//    private static final String FILE_NAME2 = "C:\\Users\\minht\\Desktop\\thuce2.xlsx";

    public static void main(String[] args) {
        ArrayList<String> listSerial = new ArrayList<String>();
        try {

            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                int i = 0;
                while (cellIterator.hasNext()) {
                    i++;
                    Cell currentCell = cellIterator.next();
//                    if (currentCell.getCellTypeEnum() == CellType.STRING && i%2 ==0) {
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        System.out.println(currentCell.getStringCellValue());
                        listSerial.add(currentCell.getStringCellValue());
                    }
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println(getDuplicate(listSerial));
    }
    private static String getDuplicate(ArrayList<String> list){
        String abs = "(";
        int t  = 0;
        for (int i = 0 ; i<list.size()-1;i++){
            try{
                if (list.get(i).equals(list.get(i+1))){
                    t ++;
                    abs = abs + list.get(i)+",";
                    i+=2;
                }
            }catch (Exception e){
                System.out.println(e);
                abs = abs +  ")";
                return abs;
            }

        }

        abs = abs +")";
        return String.valueOf(t) +abs;
    }
}
