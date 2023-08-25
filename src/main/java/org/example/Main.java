package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.entity.ExcelTable;
import org.example.entity.FData;
import org.example.entity.RowAndColumn;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Main {

    private static String fileUrl = "C:\\Users\\fu\\Desktop\\ready_convert_data.xlsx";

    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";


    List<String> list = new ArrayList<>();


    static List<ExcelTable> tableList = new ArrayList<>();

    private static void init() {

//        tableList.add(getMSKHK0301());
//        tableList.add(getMSKZY0101());
//        tableList.add(getMSKZJ0101());
//        tableList.add(getMSKTB0101());
//        tableList.add(getMSKHK0401());
//        tableList.add(getMSKGY0101());
//        tableList.add(getMSKFT0101());
//        tableList.add(getMSKHK0201());
//        tableList.add(getMSKFM0101());
//        tableList.add(getMSKMQ0101());

//        tableList.add(getMSKFK0101());
//        tableList.add(getMSKZF0101());
        tableList.add(getMSKZT0101());
    }

    // MSK-119
    private static ExcelTable getMSKZT0101() {
        ExcelTable table1 = new ExcelTable("MSK-ZT-01-01");

        FData table1status = FData.statusFData();
        RowAndColumn table1status1 = new RowAndColumn(385, 390, 0);
        RowAndColumn table1status2 = new RowAndColumn(385, 391, 5);
        table1status.setRowAndColumnList(Arrays.asList(table1status1, table1status2));

        FData table1testData = FData.testDataFData();
        RowAndColumn table1testData1 = new RowAndColumn(385, 393 , 10);
        table1testData.setRowAndColumnList(Arrays.asList(table1testData1));

        table1.setStatusDataList(Arrays.asList(table1status, table1testData));
        return table1;
    }

    // MSK-115A-PM4-CDK
    private static ExcelTable getMSKZF0101() {
        ExcelTable table1 = new ExcelTable("MSK-ZF-01-01");

        FData table1status = FData.statusFData();
        RowAndColumn table1status1 = new RowAndColumn(368, 373, 0);
        RowAndColumn table1status2 = new RowAndColumn(368, 376, 5);
        table1status.setRowAndColumnList(Arrays.asList(table1status1, table1status2));

        FData table1testData = FData.testDataFData();
        RowAndColumn table1testData1 = new RowAndColumn(368, 377, 10);
        table1testData.setRowAndColumnList(Arrays.asList(table1testData1));

        table1.setStatusDataList(Arrays.asList(table1status, table1testData));
        return table1;
    }

    // MSK-140-PM-CDK
    private static ExcelTable getMSKFK0101() {
        ExcelTable table1 = new ExcelTable("MSK-FK-01-01");

        FData table1status = FData.statusFData();
        RowAndColumn table1status1 = new RowAndColumn(292, 303, 0);
        RowAndColumn table1status2 = new RowAndColumn(292, 303, 5);
        table1status.setRowAndColumnList(Arrays.asList(table1status1, table1status2));

        FData table1testData = FData.testDataFData();
        RowAndColumn table1testData1 = new RowAndColumn(292, 307, 10);
        table1testData.setRowAndColumnList(Arrays.asList(table1testData1));

        table1.setStatusDataList(Arrays.asList(table1status, table1testData));
        return table1;
    }

    // MSK-PCA-PTM-CDK-覆膜模切叠片一体机（模切段）
    private static ExcelTable getMSKMQ0101() {
        ExcelTable table1 = new ExcelTable("MSK-MQ-01-01");

        FData table1status = FData.statusFData();
        RowAndColumn table1status1 = new RowAndColumn(232, 244, 0);
        RowAndColumn table1status2 = new RowAndColumn(232, 236, 5);
        table1status.setRowAndColumnList(Arrays.asList(table1status1, table1status2));

        FData table1testData = FData.testDataFData();
        RowAndColumn table1testData1 = new RowAndColumn(232, 241, 10);
        table1testData.setRowAndColumnList(Arrays.asList(table1testData1));

        table1.setStatusDataList(Arrays.asList(table1status, table1testData));
        return table1;
    }

    // MSK-PCA-PTM-CDK-覆膜模切叠片一体机（覆膜段）
    private static ExcelTable getMSKFM0101() {
        ExcelTable table1 = new ExcelTable("MSK-FM-01-01");

        FData table1status = FData.statusFData();
        RowAndColumn table1status1 = new RowAndColumn(216, 221, 0);
        RowAndColumn table1status2 = new RowAndColumn(216, 224, 5);
        table1status.setRowAndColumnList(Arrays.asList(table1status1, table1status2));

        FData table1testData = FData.testDataFData();
        RowAndColumn table1testData1 = new RowAndColumn(216, 228, 10);
        table1testData.setRowAndColumnList(Arrays.asList(table1testData1));

        table1.setStatusDataList(Arrays.asList(table1status, table1testData));
        return table1;
    }

    // MSK-DZF-H2190
    private static ExcelTable getMSKHK0201() {
        ExcelTable table1 = new ExcelTable("MSK-HK-02-01");

        FData table1status = FData.statusFData();
        RowAndColumn table1status1 = new RowAndColumn(194, 205, 0);
        RowAndColumn table1status2 = new RowAndColumn(194, 211, 5);
        table1status.setRowAndColumnList(Arrays.asList(table1status1, table1status2));

        FData table1testData = FData.testDataFData();
        RowAndColumn table1testData1 = new RowAndColumn(194, 211, 10);
        table1testData.setRowAndColumnList(Arrays.asList(table1testData1));

        table1.setStatusDataList(Arrays.asList(table1status, table1testData));
        return table1;
    }

    // MSK-CSC-B350-2
    private static ExcelTable getMSKFT0101() {
        ExcelTable table1 = new ExcelTable("MSK-FT-01-01");

        FData table1status = FData.statusFData();
        RowAndColumn table1status1 = new RowAndColumn(175, 184, 0);
        RowAndColumn table1status2 = new RowAndColumn(175, 187, 5);
        table1status.setRowAndColumnList(Arrays.asList(table1status1, table1status2));

        FData table1testData = FData.testDataFData();
        RowAndColumn table1testData1 = new RowAndColumn(175, 184, 10);
        table1testData.setRowAndColumnList(Arrays.asList(table1testData1));

        table1.setStatusDataList(Arrays.asList(table1status, table1testData));
        return table1;
    }

    // MSK-DPC-B320
    private static ExcelTable getMSKGY0101() {
        ExcelTable table1 = new ExcelTable("MSK-GY-01-01");

        FData table1status = FData.statusFData();
        RowAndColumn table1status1 = new RowAndColumn(162, 169, 0);
        RowAndColumn table1status2 = new RowAndColumn(162, 171, 5);
        table1status.setRowAndColumnList(Arrays.asList(table1status1, table1status2));

        FData table1testData = FData.testDataFData();
        RowAndColumn table1testData1 = new RowAndColumn(162, 171, 10);
        table1testData.setRowAndColumnList(Arrays.asList(table1testData1));

        table1.setStatusDataList(Arrays.asList(table1status, table1testData));
        return table1;
    }

    // MSK-DZF-4190
    private static ExcelTable getMSKHK0401() {
        ExcelTable table1 = new ExcelTable("MSK-HK-04-01");

        FData table1status = FData.statusFData();
        RowAndColumn table1status1 = new RowAndColumn(116, 135, 0);
        RowAndColumn table1status2 = new RowAndColumn(116, 151, 5);
        table1status.setRowAndColumnList(Arrays.asList(table1status1, table1status2));

        FData table1testData = FData.testDataFData();
        RowAndColumn table1testData1 = new RowAndColumn(116, 155, 10);
        table1testData.setRowAndColumnList(Arrays.asList(table1testData1));

        table1.setStatusDataList(Arrays.asList(table1status, table1testData));
        return table1;
    }


    // MSK-AFA-DE400L-CL5
    private static ExcelTable getMSKTB0101() {
        ExcelTable table1 = new ExcelTable("MSK-TB-01-01");

        FData table1status = FData.statusFData();
        RowAndColumn table1status1 = new RowAndColumn(81, 101, 0);
        RowAndColumn table1status2 = new RowAndColumn(81, 107, 5);
        table1status.setRowAndColumnList(Arrays.asList(table1status1, table1status2));

        FData table1testData = FData.testDataFData();
        RowAndColumn table1testData1 = new RowAndColumn(81, 106, 10);
        table1testData.setRowAndColumnList(Arrays.asList(table1testData1));

        table1.setStatusDataList(Arrays.asList(table1status, table1testData));
        return table1;
    }

    // MSK-SFM-9-10L-2
    private static ExcelTable getMSKZJ0101() {
        ExcelTable table1 = new ExcelTable("MSK-ZJ-01-01");

        FData table1status = FData.statusFData();
        RowAndColumn table1status1 = new RowAndColumn(40, 52, 0);
        RowAndColumn table1status2 = new RowAndColumn(40, 46, 5);
        table1status.setRowAndColumnList(Arrays.asList(table1status1, table1status2));

        FData table1testData = FData.testDataFData();
        RowAndColumn table1testData1 = new RowAndColumn(40, 49, 10);
        table1testData.setRowAndColumnList(Arrays.asList(table1testData1));

        table1.setStatusDataList(Arrays.asList(table1status, table1testData));
        return table1;
    }


    private static ExcelTable getMSKHK0301() {
        ExcelTable table1 = new ExcelTable("MSK-HK-03-01");

        FData table1status = FData.statusFData();
        RowAndColumn table1status1 = new RowAndColumn(4, 20, 0);
        RowAndColumn table1status2 = new RowAndColumn(4, 30, 5);
        table1status.setRowAndColumnList(Arrays.asList(table1status1, table1status2));

        FData table1testData = FData.testDataFData();
        RowAndColumn table1testData1 = new RowAndColumn(4, 35, 10);
        table1testData.setRowAndColumnList(Arrays.asList(table1testData1));

        table1.setStatusDataList(Arrays.asList(table1status, table1testData));
        return table1;
    }

    private static ExcelTable getMSKZY0101() {
        ExcelTable tableMSKzy0101 = new ExcelTable("MSK-ZY-01-01");

        FData tableMSKzy0101status = FData.statusFData();
        RowAndColumn tableMSKzy0101status1 = new RowAndColumn(319, 332, 0);
        RowAndColumn tableMSKzy0101status2 = new RowAndColumn(318, 338, 5);
        tableMSKzy0101status.setRowAndColumnList(Arrays.asList(tableMSKzy0101status1, tableMSKzy0101status2));

        FData tableMSKzy0101testData = FData.testDataFData();
        RowAndColumn tableMSKzy0101testData1 = new RowAndColumn(318, 343, 10);
        tableMSKzy0101testData.setRowAndColumnList(Arrays.asList(tableMSKzy0101testData1));

        tableMSKzy0101.setStatusDataList(Arrays.asList(tableMSKzy0101status, tableMSKzy0101testData));

        return tableMSKzy0101;
    }


    public static void main(String[] args) {
        InputStream inputStream = null;
        Workbook workbook = null;

        init();

        try {
            inputStream = new FileInputStream(fileUrl);
            String fileType = fileUrl.substring(fileUrl.lastIndexOf(".") + 1, fileUrl.length());

            workbook = getWorkbook(inputStream, fileType);

            handlerExcel(workbook, tableList);


//            parseExcel(workbook);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                inputStream.close();
                workbook.close();
            } catch (Exception e2) {
                e2.printStackTrace();
            }
        }


    }

    private static void handlerExcel(Workbook workbook, List<ExcelTable> tableList) {
        Sheet sheet = workbook.getSheetAt(0);


        for (ExcelTable excelTable : tableList) {
            for (FData fData : excelTable.getStatusDataList()) {
                List<String> filedList = new ArrayList<>();

                for (RowAndColumn rowAndColumn : fData.getRowAndColumnList()) {

                    int rowStart = rowAndColumn.getRowStart();
                    int rowEnd = rowAndColumn.getRowEnd();
                    int specifiedColumn = rowAndColumn.getSpecifiedColumn();

                    for (int i = rowStart - 1; i < rowEnd; i++) {
                        Row row = sheet.getRow(i);

                        Cell specifiedCell = row.getCell(specifiedColumn);
                        specifiedCell.setCellType(CellType.STRING);
                        String specifiedCellValue = specifiedCell.getStringCellValue();
                        specifiedCellValue = specifiedCellValue.replaceAll("\n.*", "");
                        if ("".equals(specifiedCellValue)) {
                            continue;
                        }
//                        System.err.println(specifiedCellValue);
                        filedList.add(specifiedCellValue);
                    }
                }
                fData.setFiledList(filedList);
            }
        }

        // 打印检查

        for (ExcelTable excelTable : tableList) {
            System.out.println("-------" + excelTable.getDeviceCode() + "-------");

            for (FData fData : excelTable.getStatusDataList()) {
                System.out.println(excelTable.getDeviceCode().toLowerCase().replaceAll("-", "_") + "_" + fData.getCode());
//                System.out.println("=" + fData.getCode() + "=");
                fData.getFiledList().forEach(System.out::println);
            }
        }


    }

    private static void parseExcel(Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0);

        // 自定义开始行和结束行
        int rowStart = 4;
        int rowEnd = 20;
        int specifiedColumn = 5;

        for (int i = rowStart - 1; i < rowEnd; i++) {
            Row row = sheet.getRow(i);

            Cell specifiedCell = row.getCell(specifiedColumn);
            specifiedCell.setCellType(CellType.STRING);
            String specifiedCellValue = specifiedCell.getStringCellValue();
            specifiedCellValue = specifiedCellValue.replaceAll("\n.*", "");

            System.err.println(specifiedCellValue);

        }

    }

    /**
     * 根据文件后缀名类型获取对应的工作簿对象
     *
     * @param inputStream 读取文件的输入流
     * @param fileType    文件后缀名类型（xls或xlsx）
     * @return 包含文件数据的工作簿对象
     * @throws IOException
     */
    public static Workbook getWorkbook(InputStream inputStream, String fileType) throws IOException {
        Workbook workbook = null;
        if (fileType.equalsIgnoreCase(XLS)) {
            workbook = new HSSFWorkbook(inputStream);
        } else if (fileType.equalsIgnoreCase(XLSX)) {
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook;
    }
}