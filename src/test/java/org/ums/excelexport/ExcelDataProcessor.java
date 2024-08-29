package org.ums.excelexport;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Map;

public class ExcelDataProcessor {

    public static void generateDailySummaryExcel(String rawDataSummaryExcelPath, String posConnectionTemplateExcelPath, String hierarchicalMerchantInfoExcelPath) throws IOException {
        // 获取当前日期并格式化
        SimpleDateFormat dateFormat = new SimpleDateFormat("MM.dd");
        Calendar calendar = Calendar.getInstance();
        calendar.add(Calendar.DATE, -1);  // 日期减去一天
        String formattedDate = dateFormat.format(calendar.getTime());

        // 复制并重命名POS直连参数信息模板EXCEL文件
        Path posTemplatePath = Paths.get(posConnectionTemplateExcelPath);
        Path posTargetPath = Paths.get(posTemplatePath.getParent().toString(), "POS直连参数信息模板-云南" + formattedDate + ".xlsx");
        Files.copy(posTemplatePath, posTargetPath, java.nio.file.StandardCopyOption.REPLACE_EXISTING);

        // 复制并重命名多层级商户信息EXCEL文件
        Path merchantInfoPath = Paths.get(hierarchicalMerchantInfoExcelPath);
        Path merchantTargetPath = Paths.get(merchantInfoPath.getParent().toString(), "多层级商户信息-云南" + formattedDate + ".xlsx");
        Files.copy(merchantInfoPath, merchantTargetPath, java.nio.file.StandardCopyOption.REPLACE_EXISTING);

        // 加载原始数据汇总EXCEL
        FileInputStream rawDataInputStream = new FileInputStream(rawDataSummaryExcelPath);
        Workbook rawDataWorkbook = new XSSFWorkbook(rawDataInputStream);
        Sheet rawDataSheet = rawDataWorkbook.getSheetAt(0);

        // 加载多层级商户信息EXCEL
        FileInputStream merchantInputStream = new FileInputStream(merchantTargetPath.toFile());
        Workbook merchantWorkbook = new XSSFWorkbook(merchantInputStream);
        Sheet merchantSheet = merchantWorkbook.getSheetAt(0);

        // 加载POS直连参数信息模板EXCEL
        FileInputStream posInputStream = new FileInputStream(posTargetPath.toFile());
        Workbook posWorkbook = new XSSFWorkbook(posInputStream);
        Sheet posSheet = posWorkbook.getSheetAt(0);

        // 定义城市与对应的E、F、G列映射关系
        Map<String, String[]> cityMapping = new HashMap<>();
        cityMapping.put("昆明市", new String[]{"13988220001", "101900001911014739", "昆明市"});
        cityMapping.put("曲靖市", new String[]{"13988220002", "101900001911014740", "曲靖市"});
        cityMapping.put("玉溪市", new String[]{"13988220003", "101900001911014741", "玉溪市"});
        cityMapping.put("昭通市", new String[]{"13988220004", "101900001911014742", "昭通市"});
        cityMapping.put("大理州", new String[]{"13988220005", "101900001911014743", "大理州"});
        cityMapping.put("楚雄州", new String[]{"13988220006", "101900001911014744", "楚雄州"});
        cityMapping.put("红河州", new String[]{"13988220007", "101900001911014745", "红河州"});
        cityMapping.put("普洱市", new String[]{"13988220008", "101900001911014746", "普洱市"});
        cityMapping.put("文山州", new String[]{"13988220009", "101900001911014747", "文山州"});
        cityMapping.put("德宏州", new String[]{"13988220010", "101900001911014748", "德宏州"});
        cityMapping.put("保山市", new String[]{"13988220011", "101900001911014749", "保山市"});
        cityMapping.put("迪庆州", new String[]{"13988220012", "101900001911014750", "迪庆州"});
        cityMapping.put("怒江州", new String[]{"13988220013", "101900001911014751", "怒江州"});
        cityMapping.put("西双版纳州", new String[]{"13988220014", "101900001911014752", "西双版纳州"});
        cityMapping.put("临沧市", new String[]{"13988220015", "101900001911014753", "临沧市"});
        cityMapping.put("丽江市", new String[]{"13988220016", "101900001911014754", "丽江市"});

        // 处理原始数据并更新多层级商户信息
        for (Row row : rawDataSheet) {
            if (row.getRowNum() == 0) {
                // 跳过标题行
                continue;
            }
            System.out.println("原始数据行号: " + row.getRowNum());
            if (row.getCell(1) == null || row.getCell(1).getCellType() == CellType.BLANK) {
                System.out.println("原始数据为空的行号: " + row.getRowNum());
                break;
            }

            int lastRowNum = merchantSheet.getLastRowNum();
            System.out.println("多层级商户新创建的行号: " + (lastRowNum + 1));
            Row newRow = merchantSheet.createRow(lastRowNum + 1);
            newRow.createCell(0).setCellValue("云南省家电家居换新销售平台");
            newRow.createCell(1).setCellValue("13988220000");
            newRow.createCell(2).setCellValue("101900001911014755");
            newRow.createCell(3).setCellValue("云南省家电家居换新销售平台");

            String cityName = getCellValueAsString(row, 4);
            for (Map.Entry<String, String[]> entry : cityMapping.entrySet()) {
                if (entry.getKey().contains(cityName)) {
                    newRow.createCell(4).setCellValue(entry.getValue()[0]);
                    newRow.createCell(5).setCellValue(entry.getValue()[1]);
                    newRow.createCell(6).setCellValue(entry.getValue()[2]);
                    break;
                }
            }

            String cellValueAsString = getCellValueAsString(row, 9);
            // 将科学计数法的字符串转换为 BigDecimal 以防止精度丢失
            BigDecimal bigDecimalValue = new BigDecimal(cellValueAsString);
            // 将 BigDecimal 转换为不带科学计数法的字符串
            String formattedValue = bigDecimalValue.toPlainString();
            newRow.createCell(7).setCellValue(formattedValue);
            newRow.createCell(8).setCellValue(getCellValueAsString(row, 20));
            newRow.createCell(9).setCellValue(getCellValueAsString(row, 13));
            newRow.createCell(10).setCellValue("零售");
        }

        // 处理原始数据并更新POS直连参数信息
        for (Row row : rawDataSheet) {
            if (row.getRowNum() == 0) {
                // 跳过标题行
                continue;
            }

            System.out.println("原始数据行号: " + row.getRowNum());

            if (row.getCell(1) == null || row.getCell(1).getCellType() == CellType.BLANK) {
                System.out.println("原始数据为空的行号: " + row.getRowNum());
                break;
            }

            int lastRowNum = posSheet.getLastRowNum();
            System.out.println("POS新创建的行号: " + (lastRowNum + 1));
            Row newRow = posSheet.createRow(lastRowNum + 1);
            newRow.createCell(0).setCellValue(getCellValueAsString(row, 20));
            newRow.createCell(1).setCellValue(getCellValueAsString(row, 21));
            // 获取第9列的值并转换为字符串
            String cellValueAsString = getCellValueAsString(row, 9);
            // 将科学计数法的字符串转换为 BigDecimal 以防止精度丢失
            BigDecimal bigDecimalValue = new BigDecimal(cellValueAsString);
            // 将 BigDecimal 转换为不带科学计数法的字符串
            String formattedValue = bigDecimalValue.toPlainString();
            newRow.createCell(2).setCellValue(formattedValue);
            newRow.createCell(3).setCellValue(getCellValueAsString(row, 19));
            newRow.createCell(4).setCellValue(getCellValueAsString(row, 15));
            newRow.createCell(5).setCellValue(getCellValueAsString(row, 16));
        }

        // 保存多层级商户信息
        try (FileOutputStream merchantOutputStream = new FileOutputStream(merchantTargetPath.toFile())) {
            merchantWorkbook.write(merchantOutputStream);
            System.out.println("多层级商户信息保存成功，最后行号: " + merchantSheet.getLastRowNum());
        }

        // 保存POS直连参数信息模板
        try (FileOutputStream posOutputStream = new FileOutputStream(posTargetPath.toFile())) {
            posWorkbook.write(posOutputStream);
            System.out.println("POS直连参数信息模板保存成功，最后行号: " + posSheet.getLastRowNum());
        }

        // 关闭文件流
        rawDataWorkbook.close();
        merchantWorkbook.close();
        posWorkbook.close();
    }

    private static String getCellValueAsString(Row row, int cellIndex) {
        switch (row.getCell(cellIndex).getCellType()) {
            case STRING:
                return row.getCell(cellIndex).getStringCellValue();
            case NUMERIC:
                return String.valueOf(row.getCell(cellIndex).getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(row.getCell(cellIndex).getBooleanCellValue());
            case FORMULA:
                return row.getCell(cellIndex).getCellFormula();
            case BLANK:
            default:
                return "";
        }
    }
}
