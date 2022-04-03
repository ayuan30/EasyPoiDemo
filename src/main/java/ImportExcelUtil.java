import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.text.NumberFormat;
import java.util.*;

//https://blog.51cto.com/u_9771070/2722382
public class ImportExcelUtil {

    public static void goodPrice(Map<String, String[]> map,String companyName){
        String[] cutNetProfitArr= map.get("扣非净利润(元)");
        String[] cutNetProfitRateArr= map.get("扣非净利润同比增长率");
        String[] otherDataArr= map.get(companyName);
        String[] goodPriceArr= new String[7];

        double avgCutNetProfit=0;
        double avgCutNetProfitRate=0;

        for (String s:
        cutNetProfitArr) {
            avgCutNetProfit+=Double.valueOf(s);
        }
        avgCutNetProfit=avgCutNetProfit/cutNetProfitArr.length;

        for (String s:
                cutNetProfitRateArr) {
            avgCutNetProfitRate+=Double.valueOf(s.replace("%",""));
        }
        avgCutNetProfitRate=avgCutNetProfitRate/cutNetProfitRateArr.length/100;


        goodPriceArr[0]=otherDataArr[0];
        goodPriceArr[1]=avgCutNetProfitRate+"";
        goodPriceArr[2]=otherDataArr[1];
        goodPriceArr[3]=((avgCutNetProfit*(1+Double.valueOf(otherDataArr[0]))*Double.valueOf(otherDataArr[1]))/(Double.valueOf(otherDataArr[2])*100000000))+"";
        goodPriceArr[4]=((avgCutNetProfit*(1+avgCutNetProfitRate)*Double.valueOf(otherDataArr[1]))/(Double.valueOf(otherDataArr[2])*100000000))+"";
        goodPriceArr[5]=((Double.valueOf(cutNetProfitArr[cutNetProfitArr.length-1])*(1+avgCutNetProfitRate)*Double.valueOf(otherDataArr[1]))/(Double.valueOf(otherDataArr[2])*100000000))+"";
        goodPriceArr[6]=((Double.valueOf(cutNetProfitArr[cutNetProfitArr.length-1])*(1+Double.valueOf(otherDataArr[0]))*Double.valueOf(otherDataArr[1]))/(Double.valueOf(otherDataArr[2])*100000000))+"";
        map.put(companyName+"好价格",goodPriceArr);

    }

    public static void totalAssetsIncreaseRate(Map<String, String[]> map){
        String[] totalAssetsArr= map.get("资产合计(元)");
        String[] totalAssetsRateArr=new String[totalAssetsArr.length];
        String[] totalAssetsArr2=new String[totalAssetsArr.length];
        totalAssetsRateArr[0]="0";
        NumberFormat nformat2 = NumberFormat.getNumberInstance();
        nformat2.setMaximumFractionDigits(0);
        totalAssetsArr2[0]=nformat2.format(Double.valueOf(totalAssetsArr[0]));
        for (int i = 1; i < totalAssetsArr.length; i++) {
            totalAssetsArr2[i]=nformat2.format(Double.valueOf(totalAssetsArr[i]));
            String d=totalAssetsArr[i];
            String d2=totalAssetsArr[i-1];
            totalAssetsRateArr[i]=String.valueOf((Double.valueOf(d)-Double.valueOf(d2))/Double.valueOf(d2));
        }
        map.put("总资产",totalAssetsArr2);
        map.put("总资产增长率",totalAssetsRateArr);

    }

    public static void debtRate(Map<String, String[]> map){
        String[] totalAssetsArr= map.get("资产合计(元)");
        String[] debtArr= map.get("负债合计(元)");
        String[] debtArr2= new String[debtArr.length];
        String[] debtRateArr=new String[totalAssetsArr.length];
        NumberFormat nformat2 = NumberFormat.getNumberInstance();
        nformat2.setMaximumFractionDigits(0);
        for (int i = 0; i < debtArr.length; i++) {
            Double d=Double.valueOf(debtArr[i]);
            debtArr2[i]=nformat2.format(d);
            debtRateArr[i]=String.valueOf(d/Double.valueOf(totalAssetsArr[i]));
        }
        map.put("总负债",debtArr2);
        map.put("总负债占比",debtRateArr);
    }

    /**
     * 准货币资金=货币资金(元)+交易性金融资产(元)+一年内到期的非流动资产里的理财产品+其他流动资产里的结构性存款
     * 有息负债=短期借款(元)+一年内到期的非流动负债(元)+长期借款(元)+应付债券(元)+其中：长期应付款(元)
     *
     */
    static String cashEquivalentsItems="货币资金(元)+交易性金融资产(元)+一年内到期的非流动资产里的理财产品+其他流动资产里的结构性存款";
    static String interestDebtItems="短期借款(元)+一年内到期的非流动负债(元)+长期借款(元)+应付债券(元)+其中：长期应付款(元)";
    public static void cashEquivalents(Map<String, String[]> map){
        String[] cashArr= map.get("货币资金(元)");
        String[] cashEquivalentsArr= new String[cashArr.length];
        double[] cashEquivalentsDoubleArr= new double[cashArr.length];
        String[] interestDebtArr= new String[cashArr.length];
        double[] interestDebtDoubleArr= new double[cashArr.length];

        // 遍历map中的值
        Iterator<Map.Entry<String, String[]>> it = map.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, String[]> entry = it.next();

            if(cashEquivalentsItems.indexOf(entry.getKey())>-1){
                for (int i = 0; i < entry.getValue().length; i++) {
                    if(entry.getValue()[i]==null){
                        entry.getValue()[i]="0";
                    }
                    cashEquivalentsDoubleArr[i]+=Double.valueOf(entry.getValue()[i]);
                }
            }

            if(interestDebtItems.indexOf(entry.getKey())>-1){
                for (int i = 0; i < entry.getValue().length; i++) {
                    if(entry.getValue()[i]==null){
                        entry.getValue()[i]="0";
                    }
                    interestDebtDoubleArr[i]+=Double.valueOf(entry.getValue()[i]);
                }
            }

        }

        NumberFormat nformat2 = NumberFormat.getNumberInstance();
        nformat2.setMaximumFractionDigits(2);

        String[] minusArr=new String[cashArr.length];
        for (int i = 0; i < minusArr.length; i++) {
            minusArr[i]=nformat2.format(cashEquivalentsDoubleArr[i]-interestDebtDoubleArr[i]);
        }

        for (int i = 0; i < cashEquivalentsDoubleArr.length; i++) {
            cashEquivalentsArr[i]=nformat2.format(cashEquivalentsDoubleArr[i]);
        }

        for (int i = 0; i < interestDebtDoubleArr.length; i++) {
            interestDebtArr[i]=nformat2.format(interestDebtDoubleArr[i]);
        }

        map.put("准货币资金",cashEquivalentsArr);
        map.put("有息负债",interestDebtArr);
        map.put("准货币差额",minusArr);
    }

    /**
     * 应收与预付合计=应收票据+应收账款+应收款项融资+合同资产+预付账款
     *
     */
    public static Map<String, Double[]> receivableNoteAdvancePaymentTotal(Map<String, Double[]> map){
        Double[] totalAssetsArr=map.get("总资产");
        Double[] receivableNoteAdvancePaymentTotalArr=new Double[totalAssetsArr.length];
        // 遍历map中的值
        Iterator<Map.Entry<String, Double[]>> it = map.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, Double[]> entry = it.next();

            if(receivableNotAdvancePaymentItems.indexOf(entry.getKey())>-1){
                for (int i = 0; i < entry.getValue().length; i++) {
                    if(receivableNoteAdvancePaymentTotalArr[i]==null){
                        receivableNoteAdvancePaymentTotalArr[i]=0d;
                    }
                    if(entry.getValue()[i]==null){
                        entry.getValue()[i]=0d;
                    }
                    receivableNoteAdvancePaymentTotalArr[i]+=entry.getValue()[i];
                }
            }

        }
        map.put("应收与预付合计",receivableNoteAdvancePaymentTotalArr);
        return map;
    }

    /**
     * 应付与预收合计=应付票据+应付账款+合同负债+预收账款
     * 应付预收应收预付差额=应付与预收合计-应收与预付合计
     *
     */
    static String payableNoteAdvanceReceiptsItems="其中：应付票据(元)+应付账款(元)+预收款项(元)+合同负债(元)";
    static String receivableNotAdvancePaymentItems="其中：应收票据(元)+应收账款(元)+应收款项融资(元)+合同资产(元)+预付款项(元)";
    public static void payableNoteAdvanceReceiptsTotalMinus(Map<String, String[]> map){
        String[] receivableArr=map.get("其中：应收票据(元)");
        int length=receivableArr.length;
        String[] payableNoteAdvanceReceiptsTotalArr=new String[length];
        String[] receivableNoteAdvancePaymentTotalArr=new String[length];

        double[] receivableNoteAdvancePaymentTotalDoubleArr=new double[length];
        double[] payableNoteAdvanceReceiptsTotalDoubleArr =new double[length];

        // 遍历map中的值
        Iterator<Map.Entry<String, String[]>> it = map.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, String[]> entry = it.next();

            if(receivableNotAdvancePaymentItems.indexOf(entry.getKey())>-1){
                for (int i = 0; i < entry.getValue().length; i++) {
                    if(entry.getValue()[i]==null){
                        entry.getValue()[i]="0";
                    }
                    receivableNoteAdvancePaymentTotalDoubleArr[i]+=Double.valueOf(entry.getValue()[i]);
                }
            }

            if(payableNoteAdvanceReceiptsItems.indexOf(entry.getKey())>-1){
                for (int i = 0; i < entry.getValue().length; i++) {
                    if(entry.getValue()[i]==null){
                        entry.getValue()[i]="0";
                    }
                    payableNoteAdvanceReceiptsTotalDoubleArr[i]+=Double.valueOf(entry.getValue()[i]);
                }
            }

        }

        NumberFormat nformat2 = NumberFormat.getNumberInstance();
        nformat2.setMaximumFractionDigits(2);

        String[] minusArr=new String[length];
        for (int i = 0; i < minusArr.length; i++) {
            minusArr[i]=nformat2.format(payableNoteAdvanceReceiptsTotalDoubleArr[i]- receivableNoteAdvancePaymentTotalDoubleArr[i]);
        }

        for (int i = 0; i < length; i++) {
            receivableNoteAdvancePaymentTotalArr[i]=nformat2.format(receivableNoteAdvancePaymentTotalDoubleArr[i]);
        }

        for (int i = 0; i < length; i++) {
            payableNoteAdvanceReceiptsTotalArr[i]=nformat2.format(payableNoteAdvanceReceiptsTotalDoubleArr[i]);
        }

        map.put("应收预付",receivableNoteAdvancePaymentTotalArr);
        map.put("应付预收",payableNoteAdvanceReceiptsTotalArr);
        map.put("应付差额",minusArr);
    }

    /**
     * （应收账款+合同资产）/资产总计
     *
     */
    static String receivableNoteContractAssetsItems="应收账款(元)+合同资产(元)";
    public static void receivableNoteContractAssetsRate(Map<String, String[]> map){
        String[] totalAssetsArr=map.get("资产合计(元)");
        int length=totalAssetsArr.length;
        String[] receivableNoteContractAssetsRateArr=new String[length];

        double[] receivableNoteContractAssetsDoubleArr=new double[length];

        // 遍历map中的值
        Iterator<Map.Entry<String, String[]>> it = map.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, String[]> entry = it.next();

            if(receivableNoteContractAssetsItems.indexOf(entry.getKey())>-1){
                for (int i = 0; i < entry.getValue().length; i++) {
                    if(entry.getValue()[i]==null){
                        entry.getValue()[i]="0";
                    }
                    receivableNoteContractAssetsDoubleArr[i]+=Double.valueOf(entry.getValue()[i]);
                }
            }

        }
        for (int i = 0; i < length; i++) {
            receivableNoteContractAssetsRateArr[i]=String.valueOf(receivableNoteContractAssetsDoubleArr[i]/Double.valueOf(totalAssetsArr[i]));
        }
        map.put("应收账款合同资产占比",receivableNoteContractAssetsRateArr);
    }

    /**
     * （固定资产+在建工程+工程物资）/资产总计
     *
     */
    static ArrayList<String> fixedAssetsItems=new ArrayList<>(Arrays.asList("固定资产合计(元)","在建工程合计(元)"));
    public static void fixedAssetsRate(Map<String, String[]> map){
        String[] totalAssetsArr=map.get("资产合计(元)");
        int length=totalAssetsArr.length;
        String[] fixedAssetsRateArr=new String[length];

        double[] fixedAssetsDoubleArr=new double[length];
        // 遍历map中的值
        Iterator<Map.Entry<String, String[]>> it = map.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, String[]> entry = it.next();

            if(fixedAssetsItems.contains(entry.getKey())){
                for (int i = 0; i < entry.getValue().length; i++) {
                    if(entry.getValue()[i]==null){
                        entry.getValue()[i]="0";
                    }
                    fixedAssetsDoubleArr[i]+=Double.valueOf(entry.getValue()[i]);
                }
            }

        }
        for (int i = 0; i < length; i++) {
            fixedAssetsRateArr[i]=String.valueOf(fixedAssetsDoubleArr[i]/Double.valueOf(totalAssetsArr[i]));
        }
        map.put("固定资产占比",fixedAssetsRateArr);
    }

    /**
     * 存货占比
     *
     */
    public static void inventoryRate(Map<String, String[]> map){
        String[] totalAssetsArr=map.get("资产合计(元)");
        String[] inventoryArr=map.get("存货(元)");
        String[] inventoryRateArr=new String[totalAssetsArr.length];
        for (int i = 0; i < inventoryArr.length; i++) {
            inventoryRateArr[i]=String.valueOf(Double.valueOf(inventoryArr[i])/Double.valueOf(totalAssetsArr[i]));
        }
        map.put("存货占比",inventoryRateArr);
    }

    public static void revenueIncreaseRate(Map<String, String[]> map){
        String[] revenueArr=map.get("其中：营业收入(元)");
        String[] revenueRateArr=new String[revenueArr.length];
        String[] revenueArr2=new String[revenueArr.length];
        revenueRateArr[0]="0";
        NumberFormat nformat2 = NumberFormat.getNumberInstance();
        nformat2.setMaximumFractionDigits(0);
        revenueArr2[0]=nformat2.format(Double.valueOf(revenueArr[0]));
        for (int i = 1; i < revenueArr.length; i++) {
            revenueArr2[i]=nformat2.format(Double.valueOf(revenueArr[i]));
            String d=revenueArr[i];
            String d2=revenueArr[i-1];
            revenueRateArr[i]=String.valueOf((Double.valueOf(d)-Double.valueOf(d2))/Double.valueOf(d2));
        }
        map.put("营业收入",revenueArr2);
        map.put("营业收入增长率",revenueRateArr);

    }

    public static void grossRate(Map<String, String[]> map){
        String[] revenueArr=map.get("其中：营业收入(元)");
        String[] costsArr=map.get("其中：营业成本(元)");
        Double[] grossRateDoubleArr=new Double[revenueArr.length];
        String[] grossRateArr=new String[revenueArr.length];
        String[] grossRateArr2=new String[revenueArr.length];
        grossRateArr2[0]="0";
        for (int i = 0; i < revenueArr.length; i++) {
            grossRateDoubleArr[i]=(Double.valueOf(revenueArr[i])-Double.valueOf(costsArr[i]))/Double.valueOf(revenueArr[i]);
        }

        grossRateArr[0]=String.valueOf(grossRateDoubleArr[0]);
        for (int i = 1; i < revenueArr.length; i++) {
            grossRateArr2[i]=String.valueOf((grossRateDoubleArr[i]-grossRateDoubleArr[i-1])/grossRateDoubleArr[i-1]);
            grossRateArr[i]=String.valueOf(grossRateDoubleArr[i]);
        }
        map.put("毛利率",grossRateArr);
        map.put("波动幅度",grossRateArr2);

    }

    public static void durationGrossRate(Map<String, String[]> map){
        String[] grossRateArr=map.get("毛利率");
        String[] revenueArr=map.get("其中：营业收入(元)");
        String[] costsArr1=map.get("销售费用(元)");
        String[] costsArr2=map.get("管理费用(元)");
        String[] costsArr3=map.get("研发费用(元)");
        String[] costsArr4=map.get("财务费用(元)");
        String[] durationGrossRateArr=new String[grossRateArr.length];
        for (int i = 0; i < grossRateArr.length; i++) {
            if(Double.valueOf(costsArr4[i])<0){
                durationGrossRateArr[i]=String.valueOf((Double.valueOf(costsArr1[i])+Double.valueOf(costsArr2[i])+Double.valueOf(costsArr3[i]))/Double.valueOf(revenueArr[i])/Double.valueOf(grossRateArr[i]));
            }else {
                durationGrossRateArr[i]=String.valueOf((Double.valueOf(costsArr1[i])+Double.valueOf(costsArr2[i])+Double.valueOf(costsArr3[i])+Double.valueOf(costsArr4[i]))/Double.valueOf(revenueArr[i])/Double.valueOf(grossRateArr[i]));
            }

        }
        map.put("期间费率毛利率占比",durationGrossRateArr);

    }

    public static void saleRate(Map<String, String[]> map){
        String[] revenueArr=map.get("其中：营业收入(元)");
        String[] costsArr1=map.get("销售费用(元)");
        String[] saleRateArr=new String[revenueArr.length];
        for (int i = 0; i < revenueArr.length; i++) {
            saleRateArr[i]=String.valueOf(Double.valueOf(costsArr1[i])/Double.valueOf(revenueArr[i]));
        }
        map.put("销售费用率",saleRateArr);

    }

    public static void mainBusinessRate(Map<String, String[]> map){
        String[] revenueArr=map.get("其中：营业收入(元)");
        String[] costsArr1=map.get("销售费用(元)");
        String[] costsArr2=map.get("管理费用(元)");
        String[] costsArr3=map.get("研发费用(元)");
        String[] costsArr4=map.get("财务费用(元)");
        String[] taxArr=map.get("营业税金及附加(元)");
        String[] costsArr=map.get("其中：营业成本(元)");
        String[] businessProfitArr=map.get("三、营业利润(元)");
        String[] mainBusinessRateArr=new String[revenueArr.length];
        String[] mainBusinessRateArr2=new String[revenueArr.length];
        for (int i = 0; i < revenueArr.length; i++) {
            Double mainBusiness;
            if(Double.valueOf(costsArr4[i])<0){
                mainBusiness=Double.valueOf(revenueArr[i])-(Double.valueOf(costsArr1[i])+Double.valueOf(costsArr2[i])+Double.valueOf(costsArr3[i]))-Double.valueOf(taxArr[i])-Double.valueOf(costsArr[i]);
            }else {
                mainBusiness=Double.valueOf(revenueArr[i])-(Double.valueOf(costsArr1[i])+Double.valueOf(costsArr2[i])+Double.valueOf(costsArr3[i])+Double.valueOf(costsArr4[i]))-Double.valueOf(taxArr[i])-Double.valueOf(costsArr[i]);
            }

            mainBusinessRateArr[i]=String.valueOf(mainBusiness/Double.valueOf(revenueArr[i]));
            mainBusinessRateArr2[i]=String.valueOf(mainBusiness/Double.valueOf(businessProfitArr[i]));
        }
        map.put("主营利润率",mainBusinessRateArr);
        map.put("主营利润营业利润",mainBusinessRateArr2);

    }

    public static void netProfitCashRate(Map<String, String[]> map){
        String[] businessCashArr=map.get("经营活动产生的现金流量净额(元)");
        String[] netProfitArr=map.get("五、净利润(元)");
        String[] netProfitCashRateArr=new String[businessCashArr.length];
        for (int i = 0; i < businessCashArr.length; i++) {
            netProfitCashRateArr[i]=String.valueOf(Double.valueOf(businessCashArr[i])/Double.valueOf(netProfitArr[i]));
        }
        map.put("净利润现金比率",netProfitCashRateArr);

    }

    public static void roeRate(Map<String, String[]> map){
        String[] netProfitParentArr=map.get("归属于母公司所有者的净利润(元)");
        String[] ownerEquityArr=map.get("归属于母公司所有者权益合计(元)");
        String[] roeRateArr=new String[netProfitParentArr.length];
        String[] roeRateArr2=new String[netProfitParentArr.length];
        for (int i = 0; i < netProfitParentArr.length; i++) {
            roeRateArr[i]=String.valueOf(Double.valueOf(netProfitParentArr[i])/Double.valueOf(ownerEquityArr[i]));
        }
        roeRateArr2[0]="0";
        for (int i = 1; i < netProfitParentArr.length; i++) {
            roeRateArr2[i]=String.valueOf((Double.valueOf(netProfitParentArr[i])-Double.valueOf(netProfitParentArr[i-1]))/Double.valueOf(netProfitParentArr[i-1]));
        }
        map.put("ROE",roeRateArr);
        map.put("归母净利润增长率",roeRateArr2);

    }

    public static void growRate(Map<String, String[]> map){
        String[] growArr=map.get("购建固定资产、无形资产和其他长期资产支付的现金(元)");
        String[] businessCashArr=map.get("经营活动产生的现金流量净额(元)");
        String[] growRateArr=new String[growArr.length];
        for (int i = 0; i < growArr.length; i++) {
            growRateArr[i]=String.valueOf(Double.valueOf(growArr[i])/Double.valueOf(businessCashArr[i]));
        }
        map.put("增长潜质率",growRateArr);

    }

    /**
     * 商誉占比
     *
     */
    public static void goodwillRate(Map<String, String[]> map){
        String[] totalAssetsArr=map.get("资产合计(元)");
        String[] goodwillArr=map.get("商誉(元)");
        if(goodwillArr==null){
            goodwillArr=new String[totalAssetsArr.length];
            for (int i = 0; i < goodwillArr.length; i++) {
                goodwillArr[i]="0";
            }
            map.put("商誉(元)",goodwillArr);
        }
        String[] goodwillRateArr=new String[totalAssetsArr.length];
        for (int i = 0; i < goodwillArr.length; i++) {
            goodwillRateArr[i]=String.valueOf(Double.valueOf(goodwillArr[i])/Double.valueOf(totalAssetsArr[i]));
        }
        map.put("商誉占比",goodwillRateArr);
    }

    private final static String excel2003L = ".xls"; // 2003- 版本的excel
    private final static String excel2007U = ".xlsx"; // 2007+ 版本的excel

    /**
     * 将流中的Excel数据转成List<Map>
     *
     * @param in       输入流
     * @param fileName 文件名（判断Excel版本）
     * @return
     * @throws Exception
     */
    static int endCol=8;
    public static Map<String, String[]> parseExcel(InputStream in, String fileName) throws Exception {
        // 根据文件名来创建Excel工作薄
        Workbook work = getWorkbook(in, fileName);
        if (null == work) {
            throw new Exception("创建Excel工作薄为空！");
        }
        Sheet sheet = null;
        Row row = null;
        // 返回数据
        Map<String, String[]> bdMap=new HashMap<>();
        // 遍历Excel中所有的sheet
        for (int i = 0; i < work.getNumberOfSheets(); i++) {
            sheet = work.getSheetAt(i);
            if (sheet == null) {
                continue;
            }

            // 遍历当前sheet中的所有行
            for (int j = 1; j < sheet.getLastRowNum()+5; j++) {
                row = sheet.getRow(j);
                if(row==null||row.getCell(0)==null){
                    break;
                }
                String key=row.getCell(0).getStringCellValue();
                if(fileName.contains("other_data")){
                    String[] valueArr=new String[endCol-1];
                    for (int y = 1; y <endCol-1; y++) {
                        Cell cell = row.getCell(y);
                        if(cell==null){
                            continue;
                        }
                        cell.setCellType(CellType.STRING);
                        if(row.getCell(y)==null||"".equals(row.getCell(y).getStringCellValue())||"--".equals(row.getCell(y).getStringCellValue())){
                            valueArr[y]="0";
                            continue;
                        }
                        valueArr[y-1]=row.getCell(y).getStringCellValue();
                    }
                    bdMap.put(key,valueArr);
                }else {
                    String[] valueArr=new String[endCol-1];
                    for (int y = endCol; y >1; y--) {
                        Cell cell = row.getCell(y);
                        if(cell==null){
                            continue;
                        }
                        cell.setCellType(CellType.STRING);
                        if(row.getCell(y)==null||"".equals(row.getCell(y).getStringCellValue())||"--".equals(row.getCell(y).getStringCellValue())){
                            valueArr[endCol-y]="0";
                            continue;
                        }
                        valueArr[endCol-y]=row.getCell(y).getStringCellValue();
                    }
                    bdMap.put(key,valueArr);
                }

            }
        }
        work.close();
        return bdMap;
    }

    /**
     * 投资类资产合计=以公允价值计量且其变动计入当期损益的金融资产+可供出售金融资产+其他非流动金融资产+其他权益工具投资+其他债权投资、债权投资+持有至到期投资+长期股权投资+投资性房地产
     *
     */
    static String investAssetsItems="以公允价值计量且其变动计入当期损益的金融资产+可供出售金融资产+其他非流动金融资产+其他权益工具投资+其他债权投资、债权投资+持有至到期投资+长期股权投资+投资性房地产";
    public static Map<String, Double[]> investAssetsRate(Map<String, Double[]> map){
        Double[] totalAssetsArr=map.get("总资产");
        Double[] investAssetsArr=new Double[totalAssetsArr.length];
        Double[] investAssetsRateArr=new Double[totalAssetsArr.length];
        // 遍历map中的值
        Iterator<Map.Entry<String, Double[]>> it = map.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, Double[]> entry = it.next();

            if(investAssetsItems.indexOf(entry.getKey())>-1){
                for (int i = 0; i < entry.getValue().length; i++) {
                    if(investAssetsArr[i]==null){
                        investAssetsArr[i]=0d;
                    }
                    if(entry.getValue()[i]==null){
                        entry.getValue()[i]=0d;
                    }
                    investAssetsArr[i]+=entry.getValue()[i];
                }
            }

        }
        for (int i = 0; i < investAssetsArr.length; i++) {
            investAssetsRateArr[i]=investAssetsArr[i]/totalAssetsArr[i];
        }
        map.put("投资类资产占比",investAssetsRateArr);
        return map;
    }

    public static Double nullToZero(Double n){
        if(n==null){
            return 0d;
        }else {
            return n;
        }
    }
    /**
     * 描述：根据文件后缀，自适应上传文件的版本
     *
     * @param inStr ,fileName
     * @return
     * @throws Exception
     */
    public static Workbook getWorkbook(InputStream inStr, String fileName) throws Exception {
        Workbook wb = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if (excel2003L.equals(fileType)) {
            wb = new HSSFWorkbook(inStr); // 2003-
        } else if (excel2007U.equals(fileType)) {
            wb = new XSSFWorkbook(inStr); // 2007+
        } else {
            throw new Exception("解析的文件格式有误！");
        }
        return wb;
    }

    /**
     *2019年之后
     * 与企业经营有关的资产合计=货币资金+交易性金融资产+应收票据+应收账款+应收款项融资+预付款项+存货+合同资产+长期应收款+固定资产+在建工程+使用权资产+无形资产+开发支出+长期待摊费用+递延所得税资产
     *
     * 2018年及以前
     * 与企业经营有关的资产合计=货币资金+其他流动资产里的理财产品和结构性存款+应收票据+应收账款+应收款项融资+预付款项+存货+合同资产+长期应收款+固定资产+在建工程+使用权资产+无形资产+开发支出+长期待摊费用+递延所得税资产
     *
     */
    static String businessAssetsItems="货币资金+理财产品+结构性存款+交易性金融资产+应收票据+应收账款+应收款项融资+预付款项+存货+合同资产+长期应收款+固定资产+在建工程+使用权资产+无形资产+开发支出+长期待摊费用+递延所得税资产";
    public static Map<String, Double[]> businessAssets(Map<String, Double[]> map){
        Double[] totalAssetsArr=map.get("总资产");
        Double[] businessAssetsArr=new Double[totalAssetsArr.length];
        Double[] businessAssetsRateArr=new Double[totalAssetsArr.length];
        // 遍历map中的值
        Iterator<Map.Entry<String, Double[]>> it = map.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, Double[]> entry = it.next();

            if(businessAssetsItems.indexOf(entry.getKey())>-1){
                for (int i = 0; i < entry.getValue().length; i++) {
                    if(businessAssetsArr[i]==null){
                        businessAssetsArr[i]=0d;
                    }
                    if(entry.getValue()[i]==null){
                        entry.getValue()[i]=0d;
                    }
                    businessAssetsArr[i]+=entry.getValue()[i];
                }
            }/*else {
                System.out.println(entry.getKey());
            }*/

        }
        for (int i = 0; i < businessAssetsArr.length; i++) {
            businessAssetsRateArr[i]=businessAssetsArr[i]/totalAssetsArr[i];
        }
        map.put("经营活动有关资产",businessAssetsArr);
        map.put("经营活动有关资产占比",businessAssetsRateArr);
        return map;
    }
}