package model;

import java.util.Date;

public class BDModel {
    private Date lookupDate;
    private String period;
    private String bd;
    private String year;

    public Date getLookupDate() {
        return lookupDate;
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
        return bd;
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
