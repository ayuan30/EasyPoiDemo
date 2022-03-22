import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

//https://blog.51cto.com/u_9771070/2722382
public class ImportExcelUtil {
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
    public static Map<String, Double[]> parseExcel(InputStream in, String fileName) throws Exception {
        // 根据文件名来创建Excel工作薄
        Workbook work = getWorkbook(in, fileName);
        if (null == work) {
            throw new Exception("创建Excel工作薄为空！");
        }
        Sheet sheet = null;
        Row row = null;
        // 返回数据
        Map<String, Double[]> bdMap=new HashMap<>();
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
                Double[] valueArr=new Double[row.getLastCellNum()-1];
                for (int y = 1; y < row.getLastCellNum(); y++) {
                    if(row.getCell(y)==null){
                        continue;
                    }
                    valueArr[y-1]=row.getCell(y).getNumericCellValue();
                }

                bdMap.put(key,valueArr);
            }
        }
        work.close();
        return bdMap;
    }

    public static Map<String, Double[]> totalAssetsIncreaseRate(Map<String, Double[]> map){
        Double[] totalAssetsArr=map.get("总资产");
        Double[] totalAssetsRateArr=new Double[totalAssetsArr.length];
        totalAssetsRateArr[0]=0d;
        for (int i = 1; i < totalAssetsArr.length; i++) {
            totalAssetsRateArr[i]=(totalAssetsArr[i]-totalAssetsArr[i-1])/totalAssetsArr[i-1];
        }
        map.put("总资产增长率",totalAssetsRateArr);
        return map;

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

    public static Map<String, Double[]> debtRate(Map<String, Double[]> map){
        Double[] totalAssetsArr=map.get("总资产");
        Double[] debtArr=map.get("总负债");
        Double[] debtRateArr=new Double[totalAssetsArr.length];
        for (int i = 0; i < debtArr.length; i++) {
            debtRateArr[i]=debtArr[i]/totalAssetsArr[i];
        }
        map.put("总负债占比",debtRateArr);
        return map;
    }

    /**
     * 应收与预付合计=应收票据+应收账款+应收款项融资+合同资产+预付账款
     *
     */
    static String receivableNotAdvancePaymentItems="应收票据+应收账款+应收款项融资+合同资产+预付账款";
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
    static String payableNoteAdvanceReceiptsItems="应付票据+应付账款+合同负债+预收账款";
    public static Map<String, Double[]> payableNoteAdvanceReceiptsTotalMinus(Map<String, Double[]> map){
        Double[] receivableNotAdvancePaymentTotalArr=map.get("应收与预付合计");
        Double[] payableNoteAdvanceReceiptsTotalArr=new Double[receivableNotAdvancePaymentTotalArr.length];
        Double[] minusArr=new Double[receivableNotAdvancePaymentTotalArr.length];
        // 遍历map中的值
        Iterator<Map.Entry<String, Double[]>> it = map.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, Double[]> entry = it.next();

            if(payableNoteAdvanceReceiptsItems.indexOf(entry.getKey())>-1){
                for (int i = 0; i < entry.getValue().length; i++) {
                    if(payableNoteAdvanceReceiptsTotalArr[i]==null){
                        payableNoteAdvanceReceiptsTotalArr[i]=0d;
                    }
                    if(entry.getValue()[i]==null){
                        entry.getValue()[i]=0d;
                    }
                    payableNoteAdvanceReceiptsTotalArr[i]+=entry.getValue()[i];
                }
            }

        }
        for (int i = 0; i < minusArr.length; i++) {
            minusArr[i]=payableNoteAdvanceReceiptsTotalArr[i]-receivableNotAdvancePaymentTotalArr[i];
        }
        map.put("应付与预收合计",payableNoteAdvanceReceiptsTotalArr);
        map.put("应付预收应收预付差额",minusArr);
        return map;
    }

    /**
     * （应收账款+合同资产）/资产总计
     *
     */
    static String receivableNoteContractAssetsItems="应收账款+合同资产";
    public static Map<String, Double[]> receivableNoteContractAssetsRate(Map<String, Double[]> map){
        Double[] totalAssetsArr=map.get("总资产");
        Double[] receivableNoteContractAssetsArr=new Double[totalAssetsArr.length];
        Double[] receivableNoteContractAssetsRateArr=new Double[totalAssetsArr.length];
        // 遍历map中的值
        Iterator<Map.Entry<String, Double[]>> it = map.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, Double[]> entry = it.next();

            if(receivableNoteContractAssetsItems.indexOf(entry.getKey())>-1){
                for (int i = 0; i < entry.getValue().length; i++) {
                    if(receivableNoteContractAssetsArr[i]==null){
                        receivableNoteContractAssetsArr[i]=0d;
                    }
                    if(entry.getValue()[i]==null){
                        entry.getValue()[i]=0d;
                    }
                    receivableNoteContractAssetsArr[i]+=entry.getValue()[i];
                }
            }

        }
        for (int i = 0; i < receivableNoteContractAssetsArr.length; i++) {
            receivableNoteContractAssetsRateArr[i]=receivableNoteContractAssetsArr[i]/totalAssetsArr[i];
        }
        map.put("应收账款合同资产占比",receivableNoteContractAssetsRateArr);
        return map;
    }

    /**
     * （固定资产+在建工程+工程物资）/资产总计
     *
     */
    static String fixedAssetsItems="固定资产+在建工程";
    public static Map<String, Double[]> fixedAssetsRate(Map<String, Double[]> map){
        Double[] totalAssetsArr=map.get("总资产");
        Double[] fixedAssetsArr=new Double[totalAssetsArr.length];
        Double[] fixedAssetsRateArr=new Double[totalAssetsArr.length];
        // 遍历map中的值
        Iterator<Map.Entry<String, Double[]>> it = map.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, Double[]> entry = it.next();

            if(fixedAssetsItems.indexOf(entry.getKey())>-1){
                for (int i = 0; i < entry.getValue().length; i++) {
                    if(fixedAssetsArr[i]==null){
                        fixedAssetsArr[i]=0d;
                    }
                    if(entry.getValue()[i]==null){
                        entry.getValue()[i]=0d;
                    }
                    fixedAssetsArr[i]+=entry.getValue()[i];
                }
            }

        }
        for (int i = 0; i < fixedAssetsArr.length; i++) {
            fixedAssetsRateArr[i]=fixedAssetsArr[i]/totalAssetsArr[i];
        }
        map.put("固定资产占比",fixedAssetsRateArr);
        return map;
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

    /**
     * 存货占比
     *
     */
    public static Map<String, Double[]> inventoryRate(Map<String, Double[]> map){
        Double[] totalAssetsArr=map.get("总资产");
        Double[] inventoryArr=map.get("存货");
        Double[] inventoryRateArr=new Double[totalAssetsArr.length];
        for (int i = 0; i < inventoryArr.length; i++) {
            inventoryRateArr[i]=inventoryArr[i]/totalAssetsArr[i];
        }
        map.put("存货占比",inventoryRateArr);
        return map;
    }

    /**
     * 商誉占比
     *
     */
    public static Map<String, Double[]> goodwillRate(Map<String, Double[]> map){
        Double[] totalAssetsArr=map.get("总资产");
        Double[] goodwillArr=map.get("商誉");
        Double[] goodwillRateArr=new Double[totalAssetsArr.length];
        for (int i = 0; i < goodwillArr.length; i++) {
            goodwillRateArr[i]=goodwillArr[i]/totalAssetsArr[i];
        }
        map.put("商誉占比",goodwillRateArr);
        return map;
    }

    public static Map<String, Double[]> revenueIncreaseRate(Map<String, Double[]> map){
        Double[] revenueArr=map.get("营业收入");
        Double[] revenueRateArr=new Double[revenueArr.length];
        revenueRateArr[0]=0d;
        for (int i = 1; i < revenueArr.length; i++) {
            revenueRateArr[i]=(revenueArr[i]-revenueArr[i-1])/revenueArr[i-1];
        }
        map.put("营业收入增长率",revenueRateArr);
        return map;

    }

    public static Map<String, Double[]> grossRate(Map<String, Double[]> map){
        Double[] revenueArr=map.get("营业收入");
        Double[] costsArr=map.get("营业成本");
        Double[] grossRateArr=new Double[revenueArr.length];
        for (int i = 0; i < revenueArr.length; i++) {
            grossRateArr[i]=(revenueArr[i]-costsArr[i])/revenueArr[i];
        }
        map.put("毛利率",grossRateArr);
        return map;

    }

    public static Map<String, Double[]> durationGrossRate(Map<String, Double[]> map){
        Double[] grossRateArr=map.get("毛利率");
        Double[] revenueArr=map.get("营业收入");
        Double[] costsArr1=map.get("销售费用");
        Double[] costsArr2=map.get("管理费用");
        Double[] costsArr3=map.get("研发费用");
        Double[] costsArr4=map.get("财务费用");
        Double[] durationGrossRateArr=new Double[grossRateArr.length];
        for (int i = 0; i < grossRateArr.length; i++) {
            durationGrossRateArr[i]=(nullToZero(costsArr1[i])+nullToZero(costsArr2[i])+nullToZero(costsArr3[i])+nullToZero(costsArr4[i]))/revenueArr[i]/grossRateArr[i];
        }
        map.put("期间费率毛利率占比",durationGrossRateArr);
        return map;

    }

    public static Map<String, Double[]> saleRate(Map<String, Double[]> map){
        Double[] revenueArr=map.get("营业收入");
        Double[] costsArr1=map.get("销售费用");
        Double[] saleRateArr=new Double[revenueArr.length];
        for (int i = 0; i < revenueArr.length; i++) {
            saleRateArr[i]=costsArr1[i]/revenueArr[i];
        }
        map.put("销售费用率",saleRateArr);
        return map;

    }

    public static Map<String, Double[]> mainBusinessRate(Map<String, Double[]> map){
        Double[] revenueArr=map.get("营业收入");
        Double[] costsArr1=map.get("销售费用");
        Double[] costsArr2=map.get("管理费用");
        Double[] costsArr3=map.get("研发费用");
        Double[] costsArr4=map.get("财务费用");
        Double[] taxArr=map.get("税金及附加");
        Double[] costsArr=map.get("营业成本");
        Double[] businessProfitArr=map.get("营业利润");
        Double[] mainBusinessRateArr=new Double[revenueArr.length];
        Double[] mainBusinessRateArr2=new Double[revenueArr.length];
        for (int i = 0; i < revenueArr.length; i++) {
            Double mainBusiness=revenueArr[i]-(nullToZero(costsArr1[i])+nullToZero(costsArr2[i])+nullToZero(costsArr3[i])+nullToZero(costsArr4[i]))-taxArr[i]-costsArr[i];
            mainBusinessRateArr[i]=mainBusiness/revenueArr[i];
            mainBusinessRateArr2[i]=mainBusiness/businessProfitArr[i];
        }
        map.put("主营利润率",mainBusinessRateArr);
        map.put("主营利润营业利润",mainBusinessRateArr2);
        return map;

    }

    public static Map<String, Double[]> netProfitCashRate(Map<String, Double[]> map){
        Double[] businessCashArr=map.get("经营活动产生的现金流量净额");
        Double[] netProfitArr=map.get("净利润");
        Double[] netProfitCashRateArr=new Double[businessCashArr.length];
        for (int i = 0; i < businessCashArr.length; i++) {
            netProfitCashRateArr[i]=businessCashArr[i]/netProfitArr[i];
        }
        map.put("净利润现金比率",netProfitCashRateArr);
        return map;

    }

    public static Map<String, Double[]> roeRate(Map<String, Double[]> map){
        Double[] netProfitParentArr=map.get("归属于母公司股东的净利润");
        Double[] ownerEquityArr=map.get("归属于母公司所有者权益合计");
        Double[] roeRateArr=new Double[netProfitParentArr.length];
        Double[] roeRateArr2=new Double[netProfitParentArr.length];
        for (int i = 0; i < netProfitParentArr.length; i++) {
            roeRateArr[i]=netProfitParentArr[i]/ownerEquityArr[i];
        }
        roeRateArr2[0]=0d;
        for (int i = 1; i < netProfitParentArr.length; i++) {
            roeRateArr2[i]=(netProfitParentArr[i]-netProfitParentArr[i-1])/netProfitParentArr[i-1];
        }
        map.put("ROE",roeRateArr);
        map.put("归母净利润增长率",roeRateArr2);
        return map;

    }

    public static Map<String, Double[]> growRate(Map<String, Double[]> map){
        Double[] growArr=map.get("购建固定资产、无形资产和其他长期资产支付的现金");
        Double[] businessCashArr=map.get("经营活动产生的现金流量净额");
        Double[] growRateArr=new Double[growArr.length];
        for (int i = 0; i < growArr.length; i++) {
            growRateArr[i]=growArr[i]/businessCashArr[i];
        }
        map.put("增长潜质率",growRateArr);
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

}