package service;

import model.BDModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.color.ProfileDataException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DateFormatSymbols;
import java.time.temporal.ChronoUnit;
import java.util.*;
import java.util.function.BiFunction;
import java.util.function.Predicate;

public class ExcelService {

    private static int START_BD_VALUE = -15;

    private static Predicate<Calendar> isFriday = (i) ->
            i.get(Calendar.DAY_OF_WEEK) == Calendar.FRIDAY;

    private static BiFunction<String, Calendar, Boolean> isSameYear = (month, currentPeriodStartDate) ->
        (Integer.parseInt(month) - (currentPeriodStartDate.get(Calendar.MONTH)+1)) < 0;

    private static Predicate<Cell> validCellData = (t) -> t!=null && !t.getStringCellValue().isEmpty();

    private static BiFunction<Calendar, Calendar, Boolean> checkValid = (c1, c2) -> c1!=null && c2!=null;

    private static BiFunction<Calendar, Calendar, Boolean> checkContinuity = (previousDate, currentDate) -> ChronoUnit.DAYS.between(previousDate.toInstant(), currentDate.toInstant())==1;


    public List<BDModel> extractDataFromExcel(File inputFile) {
        List<BDModel> inputScrubbedList = null;
        try {
            inputScrubbedList = new ArrayList<>();

            FileInputStream excelFile = new FileInputStream(inputFile);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet dataSheet = workbook.getSheetAt(0);

            int columnNumber = 1;
            while(columnNumber < dataSheet.getRow(0).getLastCellNum()+1) {
                Calendar currentPeriodStartDate =null;
                Calendar lastPeriodDate = Calendar.getInstance();
                int bdValue=-14;
                boolean once = false;
                for (Row r : dataSheet) {
                    if(bdValue == 0) {
                        bdValue = 1;
                    }
                    Cell c = r.getCell(columnNumber);
                    if (validCellData.test(c) && r.getRowNum() > 1) {

                        String cellDate = c.getStringCellValue();
                        if(once && checkContinuity.apply(lastPeriodDate, getCurrentDate(cellDate, currentPeriodStartDate))) {
                            Calendar temp = (Calendar) lastPeriodDate.clone();
                            temp.add(Calendar.DAY_OF_YEAR, 1);
                            bdValue = setCurrentValue(inputScrubbedList, null, bdValue, currentPeriodStartDate, lastPeriodDate, temp);
                        }
                        bdValue = setCurrentValue(inputScrubbedList, cellDate, bdValue, currentPeriodStartDate, lastPeriodDate, null);
                        once = true;
                    } else if (validCellData.test(c) && r.getRowNum() ==1){
                        String date = c.getStringCellValue();
                        setPreviousValues(inputScrubbedList, date, currentPeriodStartDate);
                    } else if (validCellData.test(c) && r.getRowNum() == 0) {
                        currentPeriodStartDate = getPeriodStartDate(c.getStringCellValue()); //Sets the Calendar to initial date of the month
                    }

                }
                if(checkValid.apply(lastPeriodDate, currentPeriodStartDate)) {
                    fillRemainingValues(lastPeriodDate, --bdValue, inputScrubbedList, String.valueOf(currentPeriodStartDate.get(Calendar.MONTH) + 1));
                }
                columnNumber++;
            }
        } catch (FileNotFoundException e) {
            System.out.println("File not found in specified path");
        } catch (IOException e) {
            System.out.println("Unable to read file");
        }
        return inputScrubbedList;
    }

    private void fillRemainingValues(Calendar lastPeriodDate, int bdValue, List<BDModel> inputScrubbedList, String period) {
        int lastDate = lastPeriodDate.getActualMaximum(Calendar.DATE);
        int date = lastPeriodDate.get(Calendar.DATE);
        while(date < lastDate ) {
            lastPeriodDate.add(Calendar.DAY_OF_YEAR, 1);
            date = lastPeriodDate.get(Calendar.DATE);
            inputScrubbedList.add(new BDModel(lastPeriodDate.getTime(), period,
                    String.valueOf(bdValue), String.valueOf(lastPeriodDate.get(Calendar.YEAR))));
        }
    }

    public Workbook createExcelWithData(List<BDModel> scrubbedList) {
        Workbook workbook = new XSSFWorkbook();
        Sheet workSheet = workbook.createSheet();
        int rowNum =1;
        intializeSheet(workSheet);
        for(BDModel bd: scrubbedList) {
            Row row = workSheet.createRow(rowNum++);
            Cell cell0 = row.createCell(0);
            Cell cell1 = row.createCell(1);
            Cell cell2= row.createCell(2);
            Cell cell3 = row.createCell(3);
            cell0.setCellValue(bd.getLookupDate());
            cell1.setCellValue(bd.getPeriod());
            cell2.setCellValue(bd.getBd());
            cell3.setCellValue(bd.getYear());
        }
        return workbook;
    }

    private void intializeSheet(Sheet workSheet) {
        Row headerRow = workSheet.createRow(0);
        Cell cell0 = headerRow.createCell(0);
        Cell cell1 = headerRow.createCell(1);
        Cell cell2= headerRow.createCell(2);
        Cell cell3 = headerRow.createCell(3);
        cell0.setCellValue("LOOKUP_DATE");
        cell1.setCellValue("PERIOD");
        cell2.setCellValue("BD");
        cell3.setCellValue("YEAR");
    }

    private int setCurrentValue(List<BDModel> inputScrubbedList, String cellDate, int bdValue, Calendar currentPeriodStartDate,Calendar lastPeriodDate, Calendar missingDate) {
        Calendar currentDate = null;
        if(cellDate == null){
           currentDate = missingDate;
       } else {
            currentDate = getCurrentDate(cellDate, currentPeriodStartDate);
        }
        inputScrubbedList.add(new BDModel(currentDate.getTime(), String.valueOf(currentPeriodStartDate.get(Calendar.MONTH)+1),
                String.valueOf(bdValue), String.valueOf(currentDate.get(Calendar.YEAR))));
        lastPeriodDate.setTime(currentDate.getTime());
        if(isFriday.test(currentDate)) {
            lastPeriodDate.setTime(insertDataForWeekends(currentDate, inputScrubbedList, bdValue, String.valueOf(currentPeriodStartDate.get(Calendar.MONTH)+1)).getTime());
        }
        return ++bdValue;
    }

    private Calendar insertDataForWeekends(Calendar currentDate, List<BDModel> inputScrubbedList, int bdValue, String period) {
        int lastDate = currentDate.getActualMaximum(Calendar.DATE);
        int date = currentDate.get(Calendar.DATE);
        if(Math.abs(date-lastDate)>1) {
            for (int i = 0; i < 2; i++) {
                currentDate.add(Calendar.DAY_OF_WEEK, 1);
                inputScrubbedList.add(new BDModel(currentDate.getTime(), period,
                        String.valueOf(bdValue), String.valueOf(currentDate.get(Calendar.YEAR))));
            }
        }else if(Math.abs(lastDate-date) == 1) {
            currentDate.add(Calendar.DAY_OF_WEEK, 1);
            inputScrubbedList.add(new BDModel(currentDate.getTime(), period,
                    String.valueOf(bdValue), String.valueOf(currentDate.get(Calendar.YEAR))));
        }
       return currentDate;
    }

    private Calendar getCurrentDate(String cellDate, Calendar currentPeriodStartDate) {
        int indexOfSlash = cellDate.indexOf("/");
        String month = cellDate.substring(0,indexOfSlash);
        String day = cellDate.substring(indexOfSlash+1);
        if(isSameYear.apply(month, currentPeriodStartDate)) {
            return new GregorianCalendar(currentPeriodStartDate.get(Calendar.YEAR)+1, Integer.parseInt(month)-1, Integer.parseInt(day.trim()));
        }
        return new GregorianCalendar(currentPeriodStartDate.get(Calendar.YEAR), Integer.parseInt(month)-1, Integer.parseInt(day.trim()));
    }

    private void setPreviousValues(List<BDModel> inputScrubbedList, String date, Calendar currentPeriodStartDate) {
        int diffInDays = getDiffDays(date, currentPeriodStartDate);
        int intialBdValue = START_BD_VALUE-diffInDays;
        int j=0;
        for (int i = intialBdValue; i<-14; i++) {
            Calendar temp = (Calendar) currentPeriodStartDate.clone(); //make a deep copy of object
            temp.add(Calendar.DAY_OF_YEAR, j++);
            inputScrubbedList.add(new BDModel(temp.getTime(), String.valueOf(currentPeriodStartDate.get(Calendar.MONTH)+1),
                    String.valueOf(i), String.valueOf(currentPeriodStartDate.get(Calendar.YEAR))));
        }
    }

    private int getDiffDays(String date, Calendar currentPeriodStartDate) {
        int indexOfSlash = date.indexOf("/");
        Calendar currentDateCalendar = new GregorianCalendar(currentPeriodStartDate.get(Calendar.YEAR), currentPeriodStartDate.get(Calendar.MONTH), Integer.parseInt(date.substring(indexOfSlash+1)));
        return  (int) ChronoUnit.DAYS.between(currentPeriodStartDate.toInstant(), currentDateCalendar.toInstant());
    }


    private Calendar getPeriodStartDate(String stringCellValue) {
        String month = stringCellValue.substring(0, 3);
        String year = stringCellValue.substring(3);
        return new GregorianCalendar(getYear(year), getMonth(month),1);
    }

    private int getMonth(String month) {
        List<String> shortMonths = Arrays.asList(new DateFormatSymbols().getShortMonths());
        for (int i =0; i<shortMonths.size(); i++) {
            if (shortMonths.get(i).equalsIgnoreCase(month)) {
                return i;
            }
        }
        return 0;
    }

    private int getYear(String year) {
        return Integer.parseInt("20"+year.trim());
    }
}

