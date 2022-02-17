package service;

import model.BDModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DateFormatSymbols;
import java.time.temporal.ChronoUnit;
import java.util.*;

public class ExcelService {

    private static int START_BD_VALUE = -15;

    public List<BDModel> extractDataFromExcel(File inputFile) {
        List<BDModel> inputScrubbedList = null;
        try {
            inputScrubbedList = new ArrayList<>();

            FileInputStream excelFile = new FileInputStream(inputFile);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet dataSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = dataSheet.iterator();
            int columnNumber = 1;
            while(columnNumber < dataSheet.getRow(0).getLastCellNum()+1) {
                Calendar currentPeriodStartDate =null;
                int bdValue=-16;
                for (Row r : dataSheet) {
                    Cell c = r.getCell(columnNumber);
                    if (c != null && r.getRowNum() > 1) {
                        String cellDate = c.getStringCellValue();
                        bdValue = setCurrentValue(inputScrubbedList, cellDate, bdValue, currentPeriodStartDate);
                    } else if (c!= null && r.getRowNum() ==1){
                        String date = c.getStringCellValue();
                        setPreviousValues(inputScrubbedList, date, currentPeriodStartDate);
                    } else if (c!=null && r.getRowNum() == 0) {
                        currentPeriodStartDate = getPeriodStartDate(c.getStringCellValue()); //Sets the Calendar to initial date of the month
                    }

                }
                columnNumber++;
            }
        } catch (FileNotFoundException e) {
            System.out.println("File not found in specified path");
        } catch (IOException e) {
            e.printStackTrace();
        }
        return inputScrubbedList;
    }

    private int setCurrentValue(List<BDModel> inputScrubbedList, String cellDate, int bdValue, Calendar currentPeriodStartDate) {
        BDModel model = new BDModel();
        model.setLookupDate(getCurrentDate(cellDate, currentPeriodStartDate).getTime());
        model.setPeriod(currentPeriodStartDate.get(Calendar.MONTH)+"");
        model.setYear(currentPeriodStartDate.get(Calendar.YEAR)+"");
        model.setBd(bdValue+"");
        inputScrubbedList.add(model);
        return bdValue++;
    }

    private Calendar getCurrentDate(String cellDate, Calendar currentPeriodStartDate) {
        int indexOfSlash = cellDate.indexOf("/");
        String month = cellDate.substring(0,indexOfSlash);
        String day = cellDate.substring(indexOfSlash+1);
        if( Integer.parseInt(month) - currentPeriodStartDate.get(Calendar.YEAR) < 0) {
            return new GregorianCalendar(currentPeriodStartDate.get(Calendar.YEAR)+1, Integer.parseInt(month), Integer.parseInt(day));
        }
        return new GregorianCalendar(currentPeriodStartDate.get(Calendar.YEAR), Integer.parseInt(month), Integer.parseInt(day));
    }

    private void setPreviousValues(List<BDModel> inputScrubbedList, String date, Calendar currentPeriodStartDate) {
        int diffInDays = getDiffDays(date, currentPeriodStartDate);
        int intialBdValue = START_BD_VALUE-diffInDays;
        for (int i = intialBdValue; i<=-15; i++) {
            BDModel model = new BDModel();
            model.setLookupDate(currentPeriodStartDate.getTime());
            model.setPeriod(currentPeriodStartDate.get(Calendar.MONTH)+"");
            model.setYear(currentPeriodStartDate.get(Calendar.YEAR)+"");
            model.setBd(i+"");
            inputScrubbedList.add(model);
        }
    }

    private int getDiffDays(String date, Calendar currentPeriodStartDate) {
        int indexOfSlash = date.indexOf("/");
        Calendar currentDateCalendar = new GregorianCalendar(currentPeriodStartDate.get(Calendar.YEAR), currentPeriodStartDate.get(Calendar.MONTH), Integer.parseInt(date.substring(indexOfSlash+1)));
        return  (int) ChronoUnit.DAYS.between(currentPeriodStartDate.toInstant(), currentDateCalendar.toInstant());
    }


    private Calendar getPeriodStartDate(String stringCellValue) {
        String month = stringCellValue.substring(0, 3);
        String year = stringCellValue.substring(4);
        return new GregorianCalendar(getYear(year), getMonth(month),1);
    }

    private int getMonth(String month) {
        List<String> shortMonths = Arrays.asList(new DateFormatSymbols().getShortMonths());
        for (int i =0; i<shortMonths.size(); i++) {
            if (shortMonths.get(i).equalsIgnoreCase(month)) {
                return i+1;
            }
        }
        return 0;
    }

    private int getYear(String year) {
        return Integer.parseInt("20"+year);
    }

}

