package model;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class BDModel {
    private Date lookupDate;
    private String period;
    private String bd;
    private String year;

    public BDModel() {
    }

    public BDModel(Date lookupDate, String period, String bd, String year) {
        this.lookupDate = lookupDate;
        this.period = period;
        this.bd = bd;
        this.year = year;
    }

    public String getLookupDate() {
        DateFormat df = new SimpleDateFormat("MM/dd/yyyy");
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

    public String getBd() {
        return "\""+bd+"BD\"";
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
}
