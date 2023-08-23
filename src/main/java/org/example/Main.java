package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class Main {

    private static String fileUrl = "C:\\Users\\fu\\Desktop\\ready_convert_data.xlsx";

    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";


    List<String> list = new ArrayList<>();

    public static void main(String[] args) {
        InputStream inputStream = null;
        Workbook workbook = null;


        try {
            inputStream = new FileInputStream(fileUrl);
            String fileType = fileUrl.substring(fileUrl.lastIndexOf(".") + 1, fileUrl.length());

            workbook = getWorkbook(inputStream, fileType);
            parseExcel(workbook);

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

    private static void parseExcel(Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0);

        int rowStart = 4;
        int rowEnd = 20;

        int specifiedColumn = 0;

        for (int i = rowStart - 1; i < rowEnd; i++) {
            Row row = sheet.getRow(i);

            Cell specifiedCell = row.getCell(specifiedColumn);
            specifiedCell.setCellType(CellType.STRING);
            String specifiedCellValue = specifiedCell.getStringCellValue();
            specifiedCellValue = specifiedCellValue.replaceAll("\n.*", "");
            System.err.println("specifiedCellValue = " + specifiedCellValue);


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