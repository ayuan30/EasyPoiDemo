import cn.afterturn.easypoi.excel.annotation.Excel;

public class GoodPrice {
    public GoodPrice() {
    }

    @Excel(name = "公司名")
    private String companyName;

    @Excel(name = "机构预测增长")
    private String orgRate;
    @Excel(name = "平均增长")
    private String avgRate;
    @Excel(name = "市盈率")
    private String ttm;
    @Excel(name = "机构预测好价格")
    private String orgPrice;
    @Excel(name = "平均增长好价格")
    private String avgPrice;
    @Excel(name = "当年平均增长好价格")
    private String currentPrice;
    @Excel(name = "当年机构预测好价格")
    private String currentPrice2;

    public String getCurrentPrice2() {
        return currentPrice2;
    }

    public void setCurrentPrice2(String currentPrice2) {
        this.currentPrice2 = currentPrice2;
    }

    public String getCurrentPrice() {
        return currentPrice;
    }

    public void setCurrentPrice(String currentPrice) {
        this.currentPrice = currentPrice;
    }

    public String getCompanyName() {
        return companyName;
    }

    public void setCompanyName(String companyName) {
        this.companyName = companyName;
    }

    public String getOrgRate() {
        return orgRate;
    }

    public void setOrgRate(String orgRate) {
        this.orgRate = orgRate;
    }

    public String getAvgRate() {
        return avgRate;
    }

    public void setAvgRate(String avgRate) {
        this.avgRate = avgRate;
    }

    public String getTtm() {
        return ttm;
    }

    public void setTtm(String ttm) {
        this.ttm = ttm;
    }

    public String getOrgPrice() {
        return orgPrice;
    }

    public void setOrgPrice(String orgPrice) {
        this.orgPrice = orgPrice;
    }

    public String getAvgPrice() {
        return avgPrice;
    }

    public void setAvgPrice(String avgPrice) {
        this.avgPrice = avgPrice;
    }
}
