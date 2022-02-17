import service.ExcelService;

import java.io.File;

public class BDConverter {

    public static void main(String args[]){
        ExcelService service = new ExcelService();
        File f = new File("D:\\BDConverter\\src\\assets\\Book1.xlsx");
        service.extractDataFromExcel(f);
    }
}
