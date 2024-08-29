package org.ums.excelexport;

import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.IOException;

@SpringBootTest
class ExcelExportApplicationTests {

    @Test
    void contextLoads() {
    }


    @Test
    void testExcelExport() throws IOException {
        String rawDataSummaryExcelPath="D://workFile/以旧换新/小U数据/商户注册导入-0828-一批-90已完成总表.xlsx";
        String posConnectionTemplateExcelPath="D://workFile/以旧换新/待生成数据/POS直连参数信息模板-云南8.27.xlsx";
        String hierarchicalMerchantInfoExcelPath="D://workFile/以旧换新/待生成数据/多层级商户信息-云南8.27.xlsx";
        ExcelDataProcessor.generateDailySummaryExcel(rawDataSummaryExcelPath,posConnectionTemplateExcelPath,hierarchicalMerchantInfoExcelPath);
    }
}
