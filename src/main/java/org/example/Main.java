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
        ExcelTable table1 = new ExcelTable("MSK-HK-03-01");

        FData table1status = FData.statusFData();
        RowAndColumn table1status1 = new RowAndColumn(4, 20, 0);
        RowAndColumn table1status2 = new RowAndColumn(4, 30, 5);
        table1status.setRowAndColumnList(Arrays.asList(table1status1, table1status2));

        FData table1testData = FData.testDataFData();
        RowAndColumn table1testData1 = new RowAndColumn(4, 35, 10);
        table1testData.setRowAndColumnList(Arrays.asList(table1testData1));

        table1.setStatusDataList(Arrays.asList(table1status, table1testData));


        tableList.add(table1);

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
//            try {
//                inputStream.close();
//                workbook.close();
//            } catch (Exception e2) {
//                e2.printStackTrace();
//            }
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