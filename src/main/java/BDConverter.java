import com.bdCalendar.model.BDModel;
import org.apache.poi.ss.usermodel.Workbook;
import com.bdCalendar.service.ExcelService;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class BDConverter {

    public static void main(String args[]){
        ExcelService service = new ExcelService();
        File f = new File("D:\\BDConverter\\src\\assets\\Book2.xlsx");
        List<BDModel> scrubbedList =  service.extractDataFromExcel(f);
        Workbook workbook = service.createExcelWithData(scrubbedList);
        try {
            FileOutputStream outputStream = new FileOutputStream("D:\\BDConverter\\src\\assets\\output.xlsx");
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
