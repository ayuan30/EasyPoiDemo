import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import org.apache.poi.ss.usermodel.Sheet;
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

    public static String[] companyArr=new String[]{"大参林","益丰","老百姓"};
    public static Map<String,Map<String, Double[]>> assetsDebtMaps=new HashMap<>();
    public static String[] itemArr=new String[]{"总资产增长率","经营活动有关资产占比","总负债占比","应付预收应收预付差额","应收账款合同资产占比","固定资产占比","投资类资产占比","存货占比","商誉占比","营业收入增长率","毛利率","期间费率毛利率占比","销售费用率","主营利润率","主营利润营业利润","净利润现金比率","ROE","归母净利润增长率","增长潜质率"};
    public static void exportExcel() throws Exception {

        for (String company:
        companyArr) {
            assetsDebtMaps.put(company,excelFileToMap(company));
        }

        // 将sheet1、sheet2、sheet3使用得map进行包装
        List<Map<String, Object>> sheetsList = new ArrayList<>();
        for (String item:
        itemArr) {
            sheetsList.add(exportMap(companyArr,item,assetsDebtMaps));
        }

        // 执行方法
        Workbook workbook = ExcelExportUtil.exportExcel(sheetsList, ExcelType.HSSF);

        FileOutputStream fos = new FileOutputStream(path+outFileName+".xlsx");
        workbook.write(fos);
        fos.close();
    }

    static String path="../Excel/";
    static String outFileName="药店行业";

    public static Map<String, Double[]> excelFileToMap(String fileName){
        String filePath = path+fileName+".xlsx";

        File file = new File(filePath);
        FileInputStream fis=null;
        try {
            fis = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        Map<String, Double[]> assetsDebtMap = null;
        try {
            assetsDebtMap = ImportExcelUtil.parseExcel(fis, file.getName());
        } catch (Exception e) {
            e.printStackTrace();
        }

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

        return assetsDebtMap;
    }

    public static Map<String, Object> exportMap(String[] companyArr,String item,Map<String,Map<String, Double[]>> assetsDebtMaps){

        ExportParams businessAssetsExportParams = new ExportParams();
        businessAssetsExportParams.setSheetName(item);
        // 创建sheet2使用得map
        Map<String, Object> businessAssetsExportMap = new HashMap<>();
        businessAssetsExportMap.put("title", businessAssetsExportParams);
        businessAssetsExportMap.put("entity", FinancialResult.class);
        List<FinancialResult> dataList=new ArrayList<>();
        for (String company:
        companyArr) {
            dataList.addAll(getList(company,item,assetsDebtMaps.get(company)));
        }
        businessAssetsExportMap.put("data", dataList);

        return businessAssetsExportMap;
    }

    public static List<FinancialResult> getList(String company,String item,Map<String, Double[]> assetsDebtMap){

        List<FinancialResult> resultList = new ArrayList<>();

            FinancialResult result = new FinancialResult();
            Double[] valueArr=assetsDebtMap.get(item);
            result.setItem(company);
            NumberFormat nformat = NumberFormat.getPercentInstance();
            nformat.setMaximumFractionDigits(2);

            int length=valueArr.length;
            for (int i = 0; i < length; i++) {
                String value=null;
                if("应付预收应收预付差额".equals(item)){
                    NumberFormat nformat2 = NumberFormat.getNumberInstance();
                    nformat2.setMaximumFractionDigits(2);
                    value=nformat2.format(valueArr[i]);
                }else {
                    value=nformat.format(valueArr[i]);
                }
                if(length==4){
                    if(i==0){
                        result.setValue17(value);
                    }else if(i==1){
                        result.setValue18(value);
                    }else if(i==2){
                        result.setValue19(value);
                    }else if(i==3){
                        result.setValue20(value);
                    }
                }else if(length==6){
                    if(i==0){
                        result.setValue15(value);
                    }else if(i==1){
                        result.setValue16(value);
                    }else if(i==2){
                        result.setValue17(value);
                    }else if(i==3){
                        result.setValue18(value);
                    }else if(i==4){
                        result.setValue19(value);
                    }else if(i==5){
                        result.setValue20(value);
                    }
                }else if(length==7){
                    if(i==0){
                        result.setValue15(value);
                    }else if(i==1){
                        result.setValue16(value);
                    }else if(i==2){
                        result.setValue17(value);
                    }else if(i==3){
                        result.setValue18(value);
                    }else if(i==4){
                        result.setValue19(value);
                    }else if(i==5){
                        result.setValue20(value);
                    }else if(i==6){
                        result.setValue21(value);
                    }
                }

            }

            resultList.add(result);

        return resultList;
    }

        static DecimalFormat df = new DecimalFormat("###,###.00");  //创建数字格式化对象
    public static void main(String[] args){
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
