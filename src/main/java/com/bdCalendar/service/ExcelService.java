package com.bdCalendar.service;

import com.bdCalendar.model.BDModel;
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
import java.util.function.BiFunction;

import java.util.function.Predicate;

public class ExcelService {

    private static final int START_BD_VALUE = -15;
    private static final BiFunction<String, Calendar, Boolean> isSameYear = (month, currentPeriodStartDate) -> (Integer.parseInt(month) - (currentPeriodStartDate.get(Calendar.MONTH) + 1)) < 0;
    private static final Predicate<Cell> validCellData = (t) -> t != null && !t.getStringCellValue().isEmpty();
    private static final BiFunction<Calendar, Calendar, Boolean> checkValid = (c1, c2) -> c1 != null && c2 != null;
    private static final BiFunction<Calendar, Calendar, Boolean> isDifferenceGreaterThanOneDay = (previousDate, currentDate) -> ChronoUnit.DAYS.between(previousDate.toInstant(), currentDate.toInstant()) > 1;

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
                for (Row r : dataSheet) {
                    if (bdValue == 0) {
                        bdValue = 1;
                    }
                    Cell c = r.getCell(columnNumber);
                    if (validCellData.test(c) && r.getRowNum() > 1) {
                        String cellDate = c.getStringCellValue();
                        /**
                         * No Saturday and Sunday checks
                         * 1. Check for difference b/w 2 dates if > 1 fill missing with previous bd value
                         */
                        if (isDifferenceGreaterThanOneDay.apply(lastPeriodDate, getCurrentDate(cellDate, currentPeriodStartDate))) {
                            Calendar temp = (Calendar) lastPeriodDate.clone();
                            temp.add(Calendar.DAY_OF_YEAR, 1);//Add one day in each iteration

                            //Execute while loop till all the missing dates are filled
                            while(temp.getTimeInMillis() < getCurrentDate(cellDate, currentPeriodStartDate).getTimeInMillis()){
                                bdValue = bdValue == 1 ? -1 : bdValue - 1;
                                bdValue = setCurrentValue(inputScrubbedList, null, bdValue, currentPeriodStartDate, lastPeriodDate, temp, true);
                                bdValue = bdValue == 0 ? 1 : bdValue;
                                temp.add(Calendar.DAY_OF_YEAR, 1);
                            }
                        }
                        bdValue = setCurrentValue(inputScrubbedList, cellDate, bdValue, currentPeriodStartDate, lastPeriodDate, null, false); //normal flow
                    } else if (validCellData.test(c) && r.getRowNum() == 1) { //extracts intial starting date in the sheet for that month
                        String date = c.getStringCellValue();
                        lastPeriodDate = getCurrentDate(date, currentPeriodStartDate);
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

    private void fillRemainingValues(Calendar lastPeriodDate, int bdValue, List<BDModel> inputScrubbedList, Calendar currentPeriodStartDate) {
        int lastDate = lastPeriodDate.getActualMaximum(Calendar.DATE);

        int date = lastPeriodDate.get(Calendar.DATE);
        int i = 1;
        while (date < lastDate) {
            lastPeriodDate.add(Calendar.DAY_OF_YEAR, 1);
            date = lastPeriodDate.get(Calendar.DATE);

            inputScrubbedList.add(new BDModel(lastPeriodDate.getTime(), String.valueOf(currentPeriodStartDate.get(Calendar.MONTH) + 1), String.valueOf(bdValue), String.valueOf(currentPeriodStartDate.get(Calendar.YEAR)), i, lastPeriodDate.getTime(), "Manual", "Manual"));

        }
    }

    public Workbook createExcelWithData(List<BDModel> scrubbedList) { //create excel output
        Workbook workbook = new XSSFWorkbook();
        Sheet workSheet = workbook.createSheet();
        int rowNum = 1;
        initializeSheet(workSheet);
        String prevDate = "";

        for (BDModel bd : scrubbedList) {
            Row row = workSheet.createRow(rowNum++);
            Cell cell0 = row.createCell(0);
            Cell cell1 = row.createCell(1);
            Cell cell2 = row.createCell(2);
            Cell cell3 = row.createCell(3);
            Cell cell4 = row.createCell(4);
            Cell cell5 = row.createCell(5);
            Cell cell6 = row.createCell(6);
            Cell cell7 = row.createCell(7);


            cell0.setCellValue(bd.getLookupDate());
            cell1.setCellValue(bd.getPeriod());
            cell2.setCellValue(bd.getBd());
            cell3.setCellValue(bd.getYear());

            // check here with previous data value
            if (prevDate.equalsIgnoreCase(bd.getLookupDate())) {
                cell4.setCellValue(bd.getSeq_No() + 1);
            } else {
                cell4.setCellValue(bd.getSeq_No());
            }

            prevDate = bd.getLookupDate();
            cell5.setCellValue(bd.getLookupDate_MMDDYYYY());
            cell6.setCellValue(bd.getCreated_By());
            cell7.setCellValue(bd.getUpdated_By());

        }
        return workbook;
    }

    private void initializeSheet(Sheet workSheet) //create headers in excel
    {
        Row headerRow = workSheet.createRow(0);
        Cell cell0 = headerRow.createCell(0);
        Cell cell1 = headerRow.createCell(1);
        Cell cell2 = headerRow.createCell(2);
        Cell cell3 = headerRow.createCell(3);
        Cell cell4 = headerRow.createCell(4);
        Cell cell5 = headerRow.createCell(5);
        Cell cell6 = headerRow.createCell(6);
        Cell cell7 = headerRow.createCell(7);


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
    private int setCurrentValue(List<BDModel> inputScrubbedList, String cellDate, int bdValue, Calendar currentPeriodStartDate, Calendar lastPeriodDate, Calendar missingDate, Boolean isManual) {
        Calendar currentDate;
        int i = 1;
        String created_by = isManual ? "Manual" : "FDAR";
        currentDate = cellDate == null ? missingDate : getCurrentDate(cellDate, currentPeriodStartDate);
        inputScrubbedList.add(new BDModel(currentDate.getTime(), String.valueOf(currentPeriodStartDate.get(Calendar.MONTH) + 1), String.valueOf(bdValue), String.valueOf(currentPeriodStartDate.get(Calendar.YEAR)), i, currentDate.getTime(), created_by, created_by));
        lastPeriodDate.setTime(currentDate.getTime());
        return ++bdValue;
    }

    //
    private Calendar getCurrentDate(String cellDate, Calendar currentPeriodStartDate) {
        int indexOfSlash = cellDate.indexOf("/");
        String month = cellDate.substring(0, indexOfSlash);
        String day = cellDate.substring(indexOfSlash + 1);
        if (isSameYear.apply(month, currentPeriodStartDate)) {
            return new GregorianCalendar(currentPeriodStartDate.get(Calendar.YEAR) + 1, Integer.parseInt(month) - 1, Integer.parseInt(day.trim()));
        }
        return new GregorianCalendar(currentPeriodStartDate.get(Calendar.YEAR), Integer.parseInt(month) - 1, Integer.parseInt(day.trim()));

    }

    //create the remaining values before startdate of given calendar
    private void setPreviousValues(List<BDModel> inputScrubbedList, String date, Calendar currentPeriodStartDate) {
        int diffInDays = getDiffDays(date, currentPeriodStartDate);
        int initialBdValue = START_BD_VALUE - diffInDays;
        int j = 0;
        int k = 1;
        for (int i = initialBdValue; i < -14; i++) {
            String created = i == START_BD_VALUE ? "FDAR" : "Manual";
            Calendar temp = (Calendar) currentPeriodStartDate.clone();
            temp.add(Calendar.DAY_OF_YEAR, j++);
            inputScrubbedList.add(new BDModel(temp.getTime(), String.valueOf(currentPeriodStartDate.get(Calendar.MONTH) + 1), String.valueOf(i), String.valueOf(currentPeriodStartDate.get(Calendar.YEAR)), k, temp.getTime(), created, created));
        }
    }

    //differenece b/w days
    private int getDiffDays(String date, Calendar currentPeriodStartDate) {
        int indexOfSlash = date.indexOf("/");
        Calendar currentDateCalendar = new GregorianCalendar(currentPeriodStartDate.get(Calendar.YEAR), currentPeriodStartDate.get(Calendar.MONTH), Integer.parseInt(date.substring(indexOfSlash + 1)));
        return (int) ChronoUnit.DAYS.between(currentPeriodStartDate.toInstant(), currentDateCalendar.toInstant());
    }

    //writes the starting date of the month
    private Calendar getPeriodStartDate(String stringCellValue) {
        String month = stringCellValue.substring(0, 3);
        String year = stringCellValue.substring(3);
        return new GregorianCalendar(getYear(year), getMonth(month), 1);
    }


    //converts the month to number
    private int getMonth(String month) {
        List<String> shortMonths = Arrays.asList(new DateFormatSymbols().getShortMonths());
        for (int i = 0; i < shortMonths.size(); i++) {
            if (shortMonths.get(i).equalsIgnoreCase(month)) {
                return i;
            }
        }
        return 0;
    }

    //writes the year in the year column by appending 20
    private int getYear(String year) {
        return Integer.parseInt("20" + year.trim());
    }
}