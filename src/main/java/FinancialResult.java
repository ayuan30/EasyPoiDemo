import cn.afterturn.easypoi.excel.annotation.Excel;

import java.math.BigDecimal;

public class FinancialResult {
    public FinancialResult() {
    }

    @Excel(name = "科目")
    private String item;

    @Excel(name = "2015")
    private String value15;
    @Excel(name = "2016")
    private String value16;
    @Excel(name = "2017")
    private String value17;
    @Excel(name = "2018")
    private String value18;
    @Excel(name = "2019")
    private String value19;
    @Excel(name = "2020")
    private String value20;
    @Excel(name = "2021")
    private String value21;

    public String getItem() {
        return item;
    }

    public void setItem(String item) {
        this.item = item;
    }

    public String getValue15() {
        return value15;
    }

    public void setValue15(String value15) {
        this.value15 = value15;
    }

    public String getValue16() {
        return value16;
    }

    public void setValue16(String value16) {
        this.value16 = value16;
    }

    public String getValue17() {
        return value17;
    }

    public void setValue17(String value17) {
        this.value17 = value17;
    }

    public String getValue18() {
        return value18;
    }

    public void setValue18(String value18) {
        this.value18 = value18;
    }

    public String getValue19() {
        return value19;
    }

    public void setValue19(String value19) {
        this.value19 = value19;
    }

    public String getValue20() {
        return value20;
    }

    public void setValue20(String value20) {
        this.value20 = value20;
    }

    public String getValue21() {
        return value21;
    }

    public void setValue21(String value21) {
        this.value21 = value21;
    }
}
