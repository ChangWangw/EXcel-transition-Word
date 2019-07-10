package com.example.poitest;

public class ExcelBean {
    private String trs1;
    private String trs2;
    private String trs3;

    @Override
    public String toString() {
        return "ExcelBean{" +
                "trs1='" + trs1 + '\'' +
                ", trs2='" + trs2 + '\'' +
                ", trs3='" + trs3 + '\'' +
                '}';
    }

    public ExcelBean(String trs1, String trs2, String trs3) {
        this.trs1 = trs1;
        this.trs2 = trs2;
        this.trs3 = trs3;
    }

    public String getTrs1() {
        return trs1;
    }

    public void setTrs1(String trs1) {
        this.trs1 = trs1;
    }

    public String getTrs2() {
        return trs2;
    }

    public void setTrs2(String trs2) {
        this.trs2 = trs2;
    }

    public String getTrs3() {
        return trs3;
    }

    public void setTrs3(String trs3) {
        this.trs3 = trs3;
    }
}
