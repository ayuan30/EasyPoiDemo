import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class AaDemo {

    static String path = "../Excel/";
    static String outFileName = "氢能";
//        static String outFileName = "持仓";
//    static String outFileName = "调味品";
//    static String outFileName = "瓷砖";
    //    static String outFileName = "药店";
//    static String outFileName = "白酒";
    public static String[] excelNameArr = new String[]{"601012", "300274", "600028", "600989", "601615", "601857", "000039", "002639", "000723", "002080", "600860", "000338", "600066", "601633", "600166"};
    public static String[] companyArr = new String[]{"隆基", "阳光电源", "中国石化", "宝丰能源", "明阳智能", "中国石油", "中集集团", "雪人股份", "美锦能源", "中材科技", "京城股份", "潍柴动力", "宇通客车", "长城汽车", "福田汽车"};
//        public static String[] excelNameArr = new String[]{"002918","603939","600887","600690","002415","002508","002475","002372","600031","600585","601318","000001"};
//    public static String[] companyArr = new String[]{"蒙娜丽莎","益丰","伊利","海尔","海康","老板","立讯","伟星","三一","海螺","平安","平安银行"};
//    public static String[] excelNameArr = new String[]{"603288","603027","600872"};
//    public static String[] companyArr = new String[]{"海天味业","千禾味业","中炬高新"};
//    public static String[] excelNameArr = new String[]{"002918"};
//    public static String[] companyArr = new String[]{"蒙娜丽莎"};
//    public static String[] excelNameArr = new String[]{"603233", "603939", "603883", "002727"};
//    public static String[] companyArr = new String[]{"大参林", "益丰", "老百姓", "一心堂"};
//    public static String[] excelNameArr = new String[]{"600519_debt_year", "000858_debt_year", "002304_debt_year"};
//    public static String[] excelNameArr = new String[]{"600519", "000858", "002304"};
//    public static String[] companyArr = new String[]{"茅台", "五粮液", "洋河"};
    public static Map<String, Map<String, String[]>> assetsDebtMaps = new HashMap<>();
    //    public static String[] itemArr=new String[]{"总资产增长率","经营活动有关资产占比","总负债占比","应付预收应收预付差额","应收账款合同资产占比","固定资产占比","投资类资产占比","存货占比","商誉占比","营业收入增长率","毛利率","期间费率毛利率占比","销售费用率","主营利润率","主营利润营业利润","净利润现金比率","ROE","归母净利润增长率","增长潜质率"};
//    public static String[] itemArr = new String[]{"总资产增长率", "总负债占比", "准货币差额","应付差额","应收账款合同资产占比","固定资产占比","存货占比"};
    public static String[] itemArr = new String[]{"总资产增长率", "总负债占比", "准货币差额", "应付差额", "应收账款合同资产占比", "固定资产占比", "存货占比", "商誉占比", "营业收入增长率", "毛利率", "期间费率毛利率占比", "销售费用率", "主营利润率", "净利润现金比率", "ROE", "增长潜质率"};

    public static void exportExcel() throws Exception {

        for (int i = 0; i < excelNameArr.length; i++) {
            assetsDebtMaps.put(companyArr[i], excelFileToMap(excelNameArr[i]));
        }

        // 将sheet1、sheet2、sheet3使用得map进行包装
        List<Map<String, Object>> sheetsList = new ArrayList<>();
        sheetsList.add(exportGoodPriceMap());
        for (String item :
                itemArr) {
            sheetsList.add(exportMap(companyArr, item, assetsDebtMaps));
        }

        // 执行方法
        Workbook workbook = ExcelExportUtil.exportExcel(sheetsList, ExcelType.HSSF);

        FileOutputStream fos = new FileOutputStream(path + outFileName + ".xlsx");
        workbook.write(fos);
        fos.close();
    }

    public static Map<String, String[]> excelFileToMap(String code) {

        Map<String, String[]> assetsDebtMap = new HashMap<>();
        String[] codeFileArr = new String[5];
        codeFileArr[0] = code + "_debt_year";
        codeFileArr[1] = code + "_benefit_year";
        codeFileArr[2] = code + "_cash_year";
        codeFileArr[3] = code + "_main_year";
        codeFileArr[4] = "other_data";
        for (int i = 0; i < codeFileArr.length; i++) {
            //        String filePath = path+fileName+".xlsx";
            String filePath = path + codeFileArr[i] + ".xls";

            File file = new File(filePath);
            FileInputStream fis = null;
            try {
                fis = new FileInputStream(file);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            try {
                assetsDebtMap.putAll(ImportExcelUtil.parseExcel(fis, file.getName()));
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        for (int i = 0; i < companyArr.length; i++) {

            //            好价格
            ImportExcelUtil.goodPrice3(assetsDebtMap, companyArr[i]);
        }
        //            总资产规模、总资产增长率
        ImportExcelUtil.totalAssetsIncreaseRate(assetsDebtMap);
        //            负债占比、负债>60%；准货币资金-有息负债<0;
        ImportExcelUtil.debtRate(assetsDebtMap);

        ImportExcelUtil.cashEquivalents(assetsDebtMap);
//            公司的竞争优势：应付预收-应收预付>0
        ImportExcelUtil.payableNoteAdvanceReceiptsTotalMinus(assetsDebtMap);
        //            产品竞争力：应收账款+合同资产占比
//            最优秀的公司（应收账款+合同资产）占总资产的比率小于 1%，优秀的公司一般小于 3%。（应收账款
//+合同资产）占总资产的比率大于 15%的公司需要淘汰掉。
        ImportExcelUtil.receivableNoteContractAssetsRate(assetsDebtMap);

        ImportExcelUtil.fixedAssetsRate(assetsDebtMap);

        //            存货占比<15%
        ImportExcelUtil.inventoryRate(assetsDebtMap);

//            商誉占总资产的比率超过 10%的公司，爆雷风险较大，需要淘汰掉。
        ImportExcelUtil.goodwillRate(assetsDebtMap);

        ImportExcelUtil.revenueIncreaseRate(assetsDebtMap);

        ImportExcelUtil.grossRate(assetsDebtMap);

        ImportExcelUtil.durationGrossRate(assetsDebtMap);

        ImportExcelUtil.saleRate(assetsDebtMap);

        ImportExcelUtil.mainBusinessRate(assetsDebtMap);
        ImportExcelUtil.netProfitCashRate(assetsDebtMap);
        ImportExcelUtil.roeRate(assetsDebtMap);
        ImportExcelUtil.growRate(assetsDebtMap);
        /*
//            与经营有关的资产及占比
        ImportExcelUtil.businessAssets(assetsDebtMap);


        ImportExcelUtil.receivableNoteAdvancePaymentTotal(assetsDebtMap);
        ImportExcelUtil.payableNoteAdvanceReceiptsTotalMinus(assetsDebtMap);

        ImportExcelUtil.receivableNoteContractAssetsRate(assetsDebtMap);
//            固定资产占比<40%

//            投资类资产占比<10%
        ImportExcelUtil.investAssetsRate(assetsDebtMap);


        ImportExcelUtil.revenueIncreaseRate(assetsDebtMap);

        ImportExcelUtil.saleRate(assetsDebtMap);



        */

        return assetsDebtMap;
    }

    public static Map<String, Object> exportGoodPriceMap() {

        ExportParams businessAssetsExportParams = new ExportParams();
        businessAssetsExportParams.setSheetName("好价格");
        // 创建sheet2使用得map
        Map<String, Object> businessAssetsExportMap = new HashMap<>();
        businessAssetsExportMap.put("title", businessAssetsExportParams);
        businessAssetsExportMap.put("entity", GoodPrice.class);
        List<GoodPrice> dataList = new ArrayList<>();
        for (String company :
                companyArr) {

            String[] goodPriceArr = assetsDebtMaps.get(company).get(company + "好价格");
            GoodPrice goodPrice = new GoodPrice();
            goodPrice.setCompanyName(company);
            NumberFormat nformat2 = NumberFormat.getNumberInstance();
            nformat2.setMaximumFractionDigits(0);

            NumberFormat nformat = NumberFormat.getPercentInstance();
            nformat.setMaximumFractionDigits(2);

            goodPrice.setMyGoodPrice(nformat2.format(Double.valueOf(goodPriceArr[0])));
            goodPrice.setMyGoodPrice2(nformat2.format(Double.valueOf(goodPriceArr[1])));
            goodPrice.setMyGoodPrice3Y(nformat2.format(Double.valueOf(goodPriceArr[2])));
            goodPrice.setMyGoodPrice3Y2(nformat2.format(Double.valueOf(goodPriceArr[3])));
            goodPrice.setCutNetProfit(nformat2.format(Double.valueOf(goodPriceArr[4]) / 100000000));
            goodPrice.setCutNetProfitRate(goodPriceArr[5]);
            goodPrice.setMyRate(nformat.format(Double.valueOf(goodPriceArr[6])));
            goodPrice.setPe(goodPriceArr[7]);

            dataList.add(goodPrice);
        }
        businessAssetsExportMap.put("data", dataList);

        return businessAssetsExportMap;
    }

    public static Map<String, Object> exportMap(String[] companyArr, String item, Map<String, Map<String, String[]>> assetsDebtMaps) {

        ExportParams businessAssetsExportParams = new ExportParams();
        businessAssetsExportParams.setSheetName(item);
        // 创建sheet2使用得map
        Map<String, Object> businessAssetsExportMap = new HashMap<>();
        businessAssetsExportMap.put("title", businessAssetsExportParams);
        businessAssetsExportMap.put("entity", FinancialResult.class);
        List<FinancialResult> dataList = new ArrayList<>();
        for (String company :
                companyArr) {
            dataList.addAll(getList(company, item, assetsDebtMaps.get(company)));
        }
        businessAssetsExportMap.put("data", dataList);

        return businessAssetsExportMap;
    }

    public static List<FinancialResult> getList(String company, String item, Map<String, String[]> assetsDebtMap) {

        List<FinancialResult> resultList = new ArrayList<>();

        if ("总资产增长率".equals(item)) {
            FinancialResult result = new FinancialResult();
            String[] valueArr = assetsDebtMap.get("总资产");
            result.setItem("总资产");
            int length = valueArr.length;
            for (int i = 0; i < length; i++) {
                String value = valueArr[i];

                setItemValue(length, i, result, value);
            }
            resultList.add(result);
        } else if ("总负债占比".equals(item)) {
            FinancialResult result = new FinancialResult();
            String[] valueArr = assetsDebtMap.get("总负债");
            result.setItem("总负债");
            int length = valueArr.length;
            for (int i = 0; i < length; i++) {
                String value = valueArr[i];

                setItemValue(length, i, result, value);
            }
            resultList.add(result);
        } else if ("准货币差额".equals(item)) {
            FinancialResult result = new FinancialResult();
            String[] valueArr = assetsDebtMap.get("准货币资金");
            result.setItem("准货币资金");
            int length = valueArr.length;
            for (int i = 0; i < length; i++) {
                String value = valueArr[i];

                setItemValue(length, i, result, value);
            }
            resultList.add(result);

            FinancialResult result2 = new FinancialResult();
            String[] valueArr2 = assetsDebtMap.get("有息负债");
            result2.setItem("有息负债");
            int length2 = valueArr.length;
            for (int i = 0; i < length2; i++) {
                String value = valueArr2[i];

                setItemValue(length2, i, result2, value);
            }
            resultList.add(result2);
        } else if ("应付差额".equals(item)) {
            FinancialResult result = new FinancialResult();
            String[] valueArr = assetsDebtMap.get("应付预收");
            result.setItem("应付预收");
            int length = valueArr.length;
            for (int i = 0; i < length; i++) {
                String value = valueArr[i];

                setItemValue(length, i, result, value);
            }
            resultList.add(result);

            FinancialResult result2 = new FinancialResult();
            String[] valueArr2 = assetsDebtMap.get("应收预付");
            result2.setItem("应收预付");
            int length2 = valueArr.length;
            for (int i = 0; i < length2; i++) {
                String value = valueArr2[i];

                setItemValue(length2, i, result2, value);
            }
            resultList.add(result2);
        } else if ("营业收入增长率".equals(item)) {
            FinancialResult result = new FinancialResult();
            String[] valueArr = assetsDebtMap.get("营业收入");
            result.setItem("营业收入");
            int length = valueArr.length;
            for (int i = 0; i < length; i++) {
                String value = valueArr[i];

                setItemValue(length, i, result, value);
            }
            resultList.add(result);
        } else if ("商誉占比".equals(item)) {
            FinancialResult result = new FinancialResult();
            String[] valueArr = assetsDebtMap.get("商誉(元)");
            result.setItem("商誉(元)");
            int length = valueArr.length;
            for (int i = 0; i < length; i++) {
                String value = valueArr[i];

                setItemValue(length, i, result, value);
            }
            resultList.add(result);
        }
        FinancialResult result = new FinancialResult();
        String[] valueArr = assetsDebtMap.get(item);
        result.setItem(company);
        NumberFormat nformat = NumberFormat.getPercentInstance();
        nformat.setMaximumFractionDigits(2);

        int length = valueArr.length;
        for (int i = 0; i < length; i++) {
            String value = valueArr[i];
            if ("准货币差额".equals(item) || "应付差额".equals(item) || "营业收入".equals(item)) {

            } else {
                value = nformat.format(Double.valueOf(value));
            }
            setItemValue(length, i, result, value);

        }
        resultList.add(result);


        if ("毛利率".equals(item)) {
            FinancialResult result2 = new FinancialResult();
            String[] valueArr2 = assetsDebtMap.get("波动幅度");
            result2.setItem("波动幅度");
            int length2 = valueArr2.length;
            for (int i = 0; i < length2; i++) {
                String value2 = valueArr2[i];
                value2 = nformat.format(Double.valueOf(value2));
                setItemValue(length2, i, result2, value2);
            }
            resultList.add(result2);
        } else if ("主营利润率".equals(item)) {
            FinancialResult result2 = new FinancialResult();
            String[] valueArr2 = assetsDebtMap.get("主营利润营业利润");
            result2.setItem("主营利润营业利润");
            int length2 = valueArr2.length;
            for (int i = 0; i < length2; i++) {
                String value2 = valueArr2[i];
                value2 = nformat.format(Double.valueOf(value2));
                setItemValue(length2, i, result2, value2);
            }
            resultList.add(result2);
        } else if ("ROE".equals(item)) {
            FinancialResult result2 = new FinancialResult();
            String[] valueArr2 = assetsDebtMap.get("归母净利润增长率");
            result2.setItem("归母净利润增长率");
            int length2 = valueArr2.length;
            for (int i = 0; i < length2; i++) {
                String value2 = valueArr2[i];
                value2 = nformat.format(Double.valueOf(value2));
                setItemValue(length2, i, result2, value2);
            }
            resultList.add(result2);
        }

        return resultList;
    }

    public static void setItemValue(int length, int i, FinancialResult result, String value) {
        if (length == 4) {
            if (i == 0) {
                result.setValue17(value);
            } else if (i == 1) {
                result.setValue18(value);
            } else if (i == 2) {
                result.setValue19(value);
            } else if (i == 3) {
                result.setValue20(value);
            }
        } else if (length == 6) {
            if (i == 0) {
                result.setValue15(value);
            } else if (i == 1) {
                result.setValue16(value);
            } else if (i == 2) {
                result.setValue17(value);
            } else if (i == 3) {
                result.setValue18(value);
            } else if (i == 4) {
                result.setValue19(value);
            } else if (i == 5) {
                result.setValue20(value);
            }
        } else if (length == 7) {
            if (i == 0) {
                result.setValue14(value);
            } else if (i == 1) {
                result.setValue15(value);
            } else if (i == 2) {
                result.setValue16(value);
            } else if (i == 3) {
                result.setValue17(value);
            } else if (i == 4) {
                result.setValue18(value);
            } else if (i == 5) {
                result.setValue19(value);
            } else if (i == 6) {
                result.setValue20(value);
            }
        } else if (length == 8) {
            if (i == 0) {
                result.setValue14(value);
            } else if (i == 1) {
                result.setValue15(value);
            } else if (i == 2) {
                result.setValue16(value);
            } else if (i == 3) {
                result.setValue17(value);
            } else if (i == 4) {
                result.setValue18(value);
            } else if (i == 5) {
                result.setValue19(value);
            } else if (i == 6) {
                result.setValue20(value);
            } else if (i == 7) {
                result.setValue21(value);
            }
        }
    }

    static DecimalFormat df = new DecimalFormat("###,###.00");  //创建数字格式化对象

    public static void main(String[] args) {
        try {
            exportExcel();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    /*{



//        String filePath = "/Users/webin/Downloads/Excel/益丰资产负债表.xlsx";
//        String filePath = "/Users/webin/Downloads/Excel/老百姓资产负债表.xlsx";
        String filePath = "/Users/webin/Downloads/Excel/大参林资产负债表.xlsx";
//        String filePath = "/Users/webin/Downloads/Excel/一心堂资产负债表.xlsx";
        File file = new File(filePath);
        FileInputStream fis=null;
        try {
            fis = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            Map<String, Double[]> assetsDebtMap = ImportExcelUtil.parseExcel(fis, file.getName());
//            总资产规模、总资产增长率
            ImportExcelUtil.totalAssetsIncreaseRate(assetsDebtMap);
//            与经营有关的资产及占比
            ImportExcelUtil.businessAssets(assetsDebtMap);
//            负债占比、负债>60%；准货币资金-有息负债<0;
            ImportExcelUtil.debtRate(assetsDebtMap);
//            公司的竞争优势：应付预收-应收预付>0
            ImportExcelUtil.receivableNoteAdvancePaymentTotal(assetsDebtMap);
            ImportExcelUtil.payableNoteAdvanceReceiptsTotalMinus(assetsDebtMap);
//            产品竞争力：应收账款+合同资产占比
//            最优秀的公司（应收账款+合同资产）占总资产的比率小于 1%，优秀的公司一般小于 3%。（应收账款
//+合同资产）占总资产的比率大于 15%的公司需要淘汰掉。
            ImportExcelUtil.receivableNoteContractAssetsRate(assetsDebtMap);
//            固定资产占比<40%
            ImportExcelUtil.fixedAssetsRate(assetsDebtMap);
//            投资类资产占比<10%
            ImportExcelUtil.investAssetsRate(assetsDebtMap);
//            存货占比<15%
            ImportExcelUtil.inventoryRate(assetsDebtMap);
//            商誉占总资产的比率超过 10%的公司，爆雷风险较大，需要淘汰掉。
            ImportExcelUtil.goodwillRate(assetsDebtMap);
            ImportExcelUtil.revenueIncreaseRate(assetsDebtMap);
            ImportExcelUtil.grossRate(assetsDebtMap);
            ImportExcelUtil.durationGrossRate(assetsDebtMap);
            ImportExcelUtil.saleRate(assetsDebtMap);
            ImportExcelUtil.mainBusinessRate(assetsDebtMap);
            ImportExcelUtil.netProfitCashRate(assetsDebtMap);
            ImportExcelUtil.roeRate(assetsDebtMap);
            ImportExcelUtil.growRate(assetsDebtMap);

            *//*for (Double val:
                    assetsDebtMap.get("总资产")) {
                NumberFormat nformat = NumberFormat.getNumberInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("导入数据一共【"+nformat.format(val)+"】行");
            }*//*

            for (Double val:
                    assetsDebtMap.get("总资产增长率")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("总资产增长率【"+nformat.format(val)+"】");
            }

            *//*for (Double val:
                    assetsDebtMap.get("经营活动有关资产")) {
                NumberFormat nformat = NumberFormat.getNumberInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("导入数据一共【"+nformat.format(val)+"】行");
            }*//*

            for (Double val:
                    assetsDebtMap.get("经营活动有关资产占比")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("经营活动有关资产占比【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("总负债占比")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("总负债占比【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("应付预收应收预付差额")) {
                NumberFormat nformat = NumberFormat.getNumberInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("应付预收应收预付差额【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("应收账款合同资产占比")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("应收账款合同资产占比【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("固定资产占比")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("固定资产占比【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("投资类资产占比")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("投资类资产占比【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("存货占比")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("存货占比【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("商誉占比")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("商誉占比【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("营业收入增长率")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("营业收入增长率【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("毛利率")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("毛利率【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("期间费率毛利率占比")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("期间费率毛利率占比【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("销售费用率")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("销售费用率【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("主营利润率")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("主营利润率【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("主营利润营业利润")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("主营利润营业利润【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("净利润现金比率")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("净利润现金比率【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("ROE")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("ROE【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("归母净利润增长率")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("归母净利润增长率【"+nformat.format(val)+"】");
            }

            for (Double val:
                    assetsDebtMap.get("增长潜质率")) {
                NumberFormat nformat = NumberFormat.getPercentInstance();
                nformat.setMaximumFractionDigits(2);

                System.out.println("增长潜质率【"+nformat.format(val)+"】");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

    }*/
}
