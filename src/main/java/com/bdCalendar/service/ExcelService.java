package com.bdCalendar.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.bdCalendar.model.BDModel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DateFormatSymbols;
import java.time.temporal.ChronoUnit;
import java.util.*;
import java.util.Calendar;
import java.util.function.BiFunction;
import java.util.function.Function;
import java.util.function.Predicate;

public class ExcelService {

    private static int START_BD_VALUE=-15;

    private static Predicate<Calendar> isFriday= (i) ->
            i.get(Calendar.DAY_OF_WEEK) == Calendar.FRIDAY;

    private static BiFunction<String, Calendar, Boolean> isSameYear = (month,currentPeriodStartDate) ->
            (Integer.parseInt(month) - (currentPeriodStartDate.get(Calendar.MONTH)+1))<0;

    private static Predicate<Cell> validCellData =(t) -> t!=null && !t.getStringCellValue().isEmpty();

    private static BiFunction<Calendar, Calendar, Boolean> checkValid= (c1,c2) -> c1!=null && c2!=null;

    private static BiFunction<Calendar, Calendar, Boolean> checkContinuity= (previousDate,currentDate) -> ChronoUnit.DAYS.between(previousDate.toInstant(), currentDate.toInstant())<=1;

    private static Function<Calendar, Boolean> checkSaturday =(date) -> date.get(Calendar.DAY_OF_WEEK)==Calendar.SATURDAY && date.get(Calendar.DATE)!=1;

    private static Function<Calendar, Boolean> checkSunday =(date) -> date.get(Calendar.DAY_OF_WEEK)==Calendar.SUNDAY;

    private static BiFunction<Calendar, Calendar, Boolean> checksameDate =(previousDate, currentDate) -> ChronoUnit.DAYS.between(previousDate.toInstant(), currentDate.toInstant())==0;


    public List<BDModel> extractDataFromExcel(File inputFile) {
        List<BDModel> inputScrubbedList = null;
        try {
            inputScrubbedList = new ArrayList<>();

            FileInputStream excelFile = new FileInputStream(inputFile);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet dataSheet = workbook.getSheetAt(0);

            int columnNumber = 1;

            while (columnNumber < dataSheet.getRow(0).getLastCellNum() + 1) {
                Calendar currentPeriodStartDate = null;
                Calendar lastPeriodDate = Calendar.getInstance();
                int bdValue = -14;
                boolean once = false;
                for (Row r : dataSheet) {
                    if (bdValue == 0) {
                        bdValue = 1;
                    }

                    Cell c = r.getCell(columnNumber);
                    if (validCellData.test(c) && r.getRowNum() > 1) {
                        String cellDate = c.getStringCellValue();
                        if (Boolean.TRUE.equals(checkSaturday.apply(getCurrentDate(cellDate, currentPeriodStartDate)))) { //check for saturday
                            if(checkContinuity.apply(lastPeriodDate, getCurrentDate(cellDate, currentPeriodStartDate))) {
                                bdValue = modifyForSaturday(inputScrubbedList, bdValue);
                            } else {
                                Calendar temp = (Calendar) lastPeriodDate.clone();
                                temp.add(Calendar.DAY_OF_YEAR, 1);
                                while(temp.getTimeInMillis() < getCurrentDate(cellDate, currentPeriodStartDate).getTimeInMillis()) {
                                    if (bdValue == 1) {
                                        bdValue = -1;

                                    } else {
                                        --bdValue;
                                    }
                                    bdValue = setCurrentValue(inputScrubbedList, null, bdValue, currentPeriodStartDate, lastPeriodDate, temp,true);
                                    if(bdValue == 0) {
                                        bdValue = 1;
                                    }
                                }

                                bdValue = modifyForSaturday(inputScrubbedList, bdValue);
                            }
                        }
                        else if (Boolean.TRUE.equals(checkSunday.apply(getCurrentDate(cellDate, currentPeriodStartDate)))) {
                            if(checkContinuity.apply(lastPeriodDate, getCurrentDate(cellDate, currentPeriodStartDate))) {
                                bdValue = modifyForSunday(inputScrubbedList, bdValue);
                            } else {
                                Calendar temp = (Calendar) lastPeriodDate.clone();
                                temp.add(Calendar.DAY_OF_YEAR, 1);
                                while(temp.getTimeInMillis() < getCurrentDate(cellDate, currentPeriodStartDate).getTimeInMillis()) {
                                    if (bdValue == 1) {
                                        bdValue = -1;

                                    } else {
                                        --bdValue;
                                    }
                                    bdValue = setCurrentValue(inputScrubbedList, null, bdValue, currentPeriodStartDate, lastPeriodDate, temp,true);
                                    if(bdValue == 0) {
                                        bdValue = 1;
                                    }
                                }

                                bdValue = modifyForSunday(inputScrubbedList, bdValue);
                            }
                        }
//                            else if (checksameDate.apply(getCurrentDate(cellDate, currentPeriodStartDate), lastPeriodDate)) {
//                            //bdValue++;
//                            modifyForSameDate(inputScrubbedList, bdValue);
//                            bdValue++;
//                        }
                        else {
                            if (once && Boolean.TRUE.equals(!checkContinuity.apply(lastPeriodDate, getCurrentDate(cellDate, currentPeriodStartDate)))) { //check for public holidays
                                Calendar temp = (Calendar) lastPeriodDate.clone();
                                temp.add(Calendar.DAY_OF_YEAR, 1);
                                if (bdValue == 1) {
                                    bdValue = -1;

                                } else {
                                    --bdValue;
                                }
                                bdValue = setCurrentValue(inputScrubbedList, null, bdValue, currentPeriodStartDate, lastPeriodDate, temp,true);
//                                if(isFriday.test(temp)) {
//                            		insertDataForWeekends(temp, inputScrubbedList, -1, currentPeriodStartDate).getTime();
//
//                            	}
                            }
                            if (bdValue == 0) {
                                bdValue = 1;
                            }
                            bdValue = setCurrentValue(inputScrubbedList, cellDate, bdValue, currentPeriodStartDate, lastPeriodDate, null,false); //normal flow
                            once = true;
                        }

                    } else if (validCellData.test(c) && r.getRowNum() == 1) { //extracts intial starting date in the sheet for that month
                        String date = c.getStringCellValue();
                        setPreviousValues(inputScrubbedList, date, currentPeriodStartDate);
                    } else if (validCellData.test(c) && r.getRowNum() == 0) { //extracts the first row of excel sheet for month and year -->Nov21
                        currentPeriodStartDate = getPeriodStartDate(c.getStringCellValue()); //sets the calendar to initial date of the month
                    }
                }
                if (Boolean.TRUE.equals(checkValid.apply(lastPeriodDate, currentPeriodStartDate))) { //fill the remaining days till month end
                    fillRemainingValues(lastPeriodDate, --bdValue, inputScrubbedList, currentPeriodStartDate);
                }
                columnNumber++;
            }
        } catch (FileNotFoundException e) {
            System.out.println("File not found in Specified patch");
        } catch (IOException e) {
            System.out.println("Unable to read file");
        }
        return inputScrubbedList;
    }

    private int modifyForSaturday(List<BDModel> inputScrubbedList,int bdValue) {
        inputScrubbedList.get(inputScrubbedList.size()-1).setBd(String.valueOf(bdValue));
        inputScrubbedList.get(inputScrubbedList.size()-2).setBd(String.valueOf(bdValue));
        inputScrubbedList.get(inputScrubbedList.size()-2).setCreated_By("FDAR");
        inputScrubbedList.get(inputScrubbedList.size()-2).setUpdated_By("FDAR");
        return ++bdValue;
    }
    //	private void modifyForSameDate(List<BDModel> inputScrubbedList,int bdValue) {
//		inputScrubbedList.get(inputScrubbedList.size()-1).setBd(String.valueOf(bdValue));
//
//
//	}
    private int modifyForSunday(List<BDModel> inputScrubbedList,int bdValue) {
        inputScrubbedList.get(inputScrubbedList.size()-1).setBd(String.valueOf(bdValue));
        inputScrubbedList.get(inputScrubbedList.size()-1).setCreated_By("FDAR");
        return ++bdValue;
    }


    private void fillRemainingValues(Calendar lastPeriodDate, int bdValue, List<BDModel> inputScrubbedList, Calendar currentPeriodStartDate) {
        int lastDate =lastPeriodDate.getActualMaximum(Calendar.DATE);

        int date = lastPeriodDate.get(Calendar.DATE);
        int i = 1;
        while (date<lastDate) {
            lastPeriodDate.add(Calendar.DAY_OF_YEAR,1);
            date = lastPeriodDate.get(Calendar.DATE);

            inputScrubbedList.add(new BDModel(lastPeriodDate.getTime(),String.valueOf(currentPeriodStartDate.get(Calendar.MONTH)+1),
                    String.valueOf(bdValue), String.valueOf(currentPeriodStartDate.get(Calendar.YEAR)),i, lastPeriodDate.getTime(),"Manual","Manual"));

        }
    }

    public Workbook createExcelWithData(List<BDModel> scrubbedList) { //create excel output
        Workbook workbook =new XSSFWorkbook();
        Sheet workSheet =workbook.createSheet();
        int rowNum=1;
        initializeSheet(workSheet);
        String prevDate =  "";

        for(BDModel bd:scrubbedList) {
            Row row =workSheet.createRow(rowNum++);
            Cell cell0=row.createCell(0);
            Cell cell1=row.createCell(1);
            Cell cell2=row.createCell(2);
            Cell cell3=row.createCell(3);
            Cell cell4=row.createCell(4);
            Cell cell5=row.createCell(5);
            Cell cell6=row.createCell(6);
            Cell cell7=row.createCell(7);


            cell0.setCellValue(bd.getLookupDate());
            cell1.setCellValue(bd.getPeriod());
            cell2.setCellValue(bd.getBd());
            cell3.setCellValue(bd.getYear());

            // check here with previous data value
            if ( prevDate.equalsIgnoreCase(bd.getLookupDate().toString())) {
                cell4.setCellValue(bd.getSeq_No()+1);
            }

            else {
                cell4.setCellValue(bd.getSeq_No());
            }

            prevDate = bd.getLookupDate().toString();
            cell5.setCellValue(bd.getLookupDate_MMDDYYYY());
            cell6.setCellValue(bd.getCreated_By());
            cell7.setCellValue(bd.getUpdated_By());

        }
        return workbook;
    }

    private void  initializeSheet(Sheet workSheet) //create headers in excel
    {
        Row headerRow =workSheet.createRow(0);
        Cell cell0 =headerRow.createCell(0);
        Cell cell1 =headerRow.createCell(1);
        Cell cell2 =headerRow.createCell(2);
        Cell cell3 =headerRow.createCell(3);
        Cell cell4 =headerRow.createCell(4);
        Cell cell5 =headerRow.createCell(5);
        Cell cell6 =headerRow.createCell(6);
        Cell cell7 =headerRow.createCell(7);


        cell0.setCellValue("LOOKUP_DATE");
        cell1.setCellValue("PERIOD");
        cell2.setCellValue("BD");
        cell3.setCellValue("YEAR");
        cell4.setCellValue("SEQ_NO");
        cell5.setCellValue("LOOKUP_DATE_MMDDYYYY");
        cell6.setCellValue("CREATED_BY");
        cell7.setCellValue("UPDATED_BY");

    }
    //
    private int setCurrentValue(List<BDModel> inputScrubbedList, String cellDate, int bdValue,Calendar currentPeriodStartDate, Calendar lastPeriodDate, Calendar missingDate, Boolean isManual)
    {
        Calendar currentDate=null;
        int i=1;
        String created_by= isManual?"Manual" : "FDAR";
        if(cellDate==null) {
            currentDate=missingDate;
        } else {

            currentDate = getCurrentDate(cellDate, currentPeriodStartDate);
        }
        inputScrubbedList.add(new BDModel(currentDate.getTime(), String.valueOf(currentPeriodStartDate.get(Calendar.MONTH)+1),
                String.valueOf(bdValue), String.valueOf(currentPeriodStartDate.get(Calendar.YEAR)),i,currentDate.getTime(),created_by,created_by));

        lastPeriodDate.setTime(currentDate.getTime());
        if(isFriday.test(currentDate)) {
            lastPeriodDate.setTime(insertDataForWeekends(currentDate, inputScrubbedList, bdValue, currentPeriodStartDate).getTime());

        }
        return ++bdValue;

    }

    //insert data for weekends
    private Calendar insertDataForWeekends(Calendar currentDate, List<BDModel> inputScrubbedList,int bdValue, Calendar currentPeriodStartDate )
    {
        int lastDate =currentDate.getActualMaximum(Calendar.DATE);
        int date =currentDate.get(Calendar.DATE);
        int j=1;
        if(Math.abs(date-lastDate)>1 || (bdValue<0 && (Math.abs(date-lastDate)==0))) {
//	if(Math.abs(date-lastDate)>1) {

            for(int i=0;i<2;i++)
            {
                currentDate.add(Calendar.DAY_OF_WEEK, 1);
                inputScrubbedList.add(new BDModel(currentDate.getTime(),String.valueOf(currentPeriodStartDate.get(Calendar.MONTH)+1),
                        String.valueOf(bdValue), String.valueOf(currentPeriodStartDate.get(Calendar.YEAR)),j,currentDate.getTime(),"Manual","Manual"));

            }
        } else if (Math.abs(lastDate-date)==1) {
            currentDate.add(Calendar.DAY_OF_WEEK, 1);
            inputScrubbedList.add(new BDModel(currentDate.getTime(),String.valueOf(currentPeriodStartDate.get(Calendar.MONTH)+1),
                    String.valueOf(bdValue), String.valueOf(currentPeriodStartDate.get(Calendar.YEAR)),j,currentDate.getTime(),"Manual","Manual"));

        }
        return currentDate;
    }

    //
    private Calendar getCurrentDate(String cellDate, Calendar currentPeriodStartDate)
    {
        int indexOfSlash =cellDate.indexOf("/");
        String month = cellDate.substring(0,indexOfSlash);
        String day = cellDate.substring(indexOfSlash+1);
        if(isSameYear.apply(month, currentPeriodStartDate))
        {
            return new GregorianCalendar(currentPeriodStartDate.get(Calendar.YEAR)+1, Integer.parseInt(month)-1, Integer.parseInt(day.trim()));
        }
        return new GregorianCalendar(currentPeriodStartDate.get(Calendar.YEAR), Integer.parseInt(month)-1, Integer.parseInt(day.trim()));

    }

    //create the remaining values before startdate of given calendar
    private void setPreviousValues(List<BDModel> inputScrubbedList, String date, Calendar currentPeriodStartDate)
    {
        int diffInDays = getDiffDays(date, currentPeriodStartDate);
        int initialBdValue = START_BD_VALUE-diffInDays;
        int j=0;
        int k=1;
        for(int i = initialBdValue; i<-14; i++)
        {
            String created = i==START_BD_VALUE?"FDAR" : "Manual";
            Calendar temp = (Calendar) currentPeriodStartDate.clone();
            temp.add(Calendar.DAY_OF_YEAR, j++);
            inputScrubbedList.add(new BDModel(temp.getTime(), String.valueOf(currentPeriodStartDate.get(Calendar.MONTH)+1),
                    String.valueOf(i), String.valueOf(currentPeriodStartDate.get(Calendar.YEAR)),k,temp.getTime(),created,created));
        } Calendar currentDate=getCurrentDate(date,currentPeriodStartDate);
        if(isFriday.test(currentDate)) {
            insertDataForWeekends(currentDate, inputScrubbedList, START_BD_VALUE, currentPeriodStartDate).getTime();

        }
        else if (Boolean.TRUE.equals(checkSaturday.apply(getCurrentDate(date, currentPeriodStartDate)))) {
            currentDate.add(Calendar.DAY_OF_WEEK, 1);
            inputScrubbedList.add(new BDModel(currentDate.getTime(),String.valueOf(currentPeriodStartDate.get(Calendar.MONTH)+1),
                    String.valueOf(START_BD_VALUE), String.valueOf(currentPeriodStartDate.get(Calendar.YEAR)),1,currentDate.getTime(),"Manual","Manual"));
        }


    }

    //differenece b/w days
    private int getDiffDays(String date, Calendar currentPeriodStartDate) {

        int indexOfSlash = date.indexOf("/");
        Calendar currentDateCalendar = new GregorianCalendar(currentPeriodStartDate.get(Calendar.YEAR),currentPeriodStartDate.get(Calendar.MONTH), Integer.parseInt(date.substring(indexOfSlash+1)));
        return (int) ChronoUnit.DAYS.between(currentPeriodStartDate.toInstant(), currentDateCalendar.toInstant());
    }

    //writes the starting date of the month
    private Calendar getPeriodStartDate(String stringCellValue) {
        String month = stringCellValue.substring(0,3);
        String year = stringCellValue.substring(3);
        return new GregorianCalendar(getYear(year),getMonth(month),1);
    }


    //converts the month to number
    private int getMonth(String month) {
        List<String> shortMonths = Arrays.asList(new  DateFormatSymbols().getShortMonths());
        for (int i=0;i<shortMonths.size();i++) {
            if (shortMonths.get(i).equalsIgnoreCase(month)) {
                return i;
            }
        }
        return 0;
    }

    //writes the year in the year column by appending 20
    private int getYear(String year) {
        return Integer.parseInt("20"+year.trim());
    }
}