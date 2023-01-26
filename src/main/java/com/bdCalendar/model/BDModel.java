package com.bdCalendar.model;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class BDModel {

    private Date lookupDate;
    private String period;
    private String bd;
    private String year;
    private int seq_No;
    private Date lookupDate_MMDDYYYY;
    private String created_By;
    private String updated_By;



    public String getCreated_By() {
        return created_By;
    }

    public void setCreated_By(String created_By) {
        this.created_By = created_By;
    }

    public String getUpdated_By() {
        return updated_By;
    }

    public void setUpdated_By(String updated_By) {
        this.updated_By = updated_By;
    }


    public int getSeq_No() {
        return seq_No;
    }

    public void setSeq_No(int seq_No) {
        this.seq_No = seq_No;
    }


    public String getLookupDate_MMDDYYYY() {
        DateFormat df=new SimpleDateFormat("MM/dd/yyyy");
        return df.format(lookupDate_MMDDYYYY);
    }

    public void setLookupDate_MMDDYYYY(Date lookupDate_MMDDYYYY) {
        this.lookupDate_MMDDYYYY = lookupDate_MMDDYYYY;
    }

    public BDModel() {

    }

    public String getLookupDate() {
        //DateFormat df=new SimpleDateFormat("dd/mm/yyyy");
        DateFormat df=new SimpleDateFormat("M/d");
        return df.format(lookupDate);
    }

    public void setLookupDate(Date lookupDate) {
        this.lookupDate = lookupDate;
    }

    public String getPeriod() {
        return period;
    }

    public void setPeriod(String period) {
        this.period = period;
    }

    public String getBd() { //"\""+bd+"BD\""
        if (Integer.parseInt(bd)>0) {
            return "BD"+bd;
        } else {
            return "-BD"+Math.abs(Integer.parseInt(bd));
        }
    }


    public void setBd(String bd) {
        this.bd = bd;
    }

    public String getYear() {
        return year;
    }

    public void setYear(String year) {
        this.year = year;
    }

    public BDModel(Date lookupDate, String period, String bd, String year, int seq_No, Date lookupDate_MMDDYYYY,
                   String created_By, String updated_By) {
        super();
        this.lookupDate = lookupDate;
        this.period = period;
        this.bd = bd;
        this.year = year;
        this.seq_No = seq_No;
        this.lookupDate_MMDDYYYY = lookupDate_MMDDYYYY;
        this.created_By = created_By;
        this.updated_By = updated_By;
    }
}