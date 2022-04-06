import cn.afterturn.easypoi.excel.annotation.Excel;

public class GoodPrice {
    public GoodPrice() {
    }

    @Excel(name = "公司名")
    private String companyName;

    @Excel(name = "我的好价格9")
    private String myGoodPrice2;
    @Excel(name = "我的好价格9（看3年）")
    private String myGoodPrice3Y2;
    @Excel(name = "我的好价格7")
    private String myGoodPrice;
    @Excel(name = "我的好价格7（看3年）")
    private String myGoodPrice3Y;
    @Excel(name = "扣非净利润")
    private String cutNetProfit;
    @Excel(name = "扣非净利润增长率")
    private String cutNetProfitRate;
    @Excel(name = "我的增长率")
    private String myRate;
    @Excel(name = "市盈率")
    private String pe;

    public String getMyGoodPrice3Y2() {
        return myGoodPrice3Y2;
    }

    public void setMyGoodPrice3Y2(String myGoodPrice3Y2) {
        this.myGoodPrice3Y2 = myGoodPrice3Y2;
    }

    public String getMyGoodPrice3Y() {
        return myGoodPrice3Y;
    }

    public void setMyGoodPrice3Y(String myGoodPrice3Y) {
        this.myGoodPrice3Y = myGoodPrice3Y;
    }

    public String getCompanyName() {
        return companyName;
    }

    public void setCompanyName(String companyName) {
        this.companyName = companyName;
    }

    public String getMyGoodPrice() {
        return myGoodPrice;
    }

    public void setMyGoodPrice(String myGoodPrice) {
        this.myGoodPrice = myGoodPrice;
    }

    public String getMyGoodPrice2() {
        return myGoodPrice2;
    }

    public void setMyGoodPrice2(String myGoodPrice2) {
        this.myGoodPrice2 = myGoodPrice2;
    }

    public String getCutNetProfit() {
        return cutNetProfit;
    }

    public void setCutNetProfit(String cutNetProfit) {
        this.cutNetProfit = cutNetProfit;
    }

    public String getCutNetProfitRate() {
        return cutNetProfitRate;
    }

    public void setCutNetProfitRate(String cutNetProfitRate) {
        this.cutNetProfitRate = cutNetProfitRate;
    }

    public String getMyRate() {
        return myRate;
    }

    public void setMyRate(String myRate) {
        this.myRate = myRate;
    }

    public String getPe() {
        return pe;
    }

    public void setPe(String pe) {
        this.pe = pe;
    }
}
