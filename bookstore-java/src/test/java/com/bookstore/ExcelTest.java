package com.bookstore;

import cn.hutool.core.date.DateTime;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.util.Date;


/**
 * @author sky
 * @date 2021/11/5 10:26
 */

public class ExcelTest {
    public static String PATH = "D:\\";

    @Disabled
    @Test
    public void cellTypeTest() throws Exception {
        FileInputStream inputStream = new FileInputStream(PATH + "xx.xls");//文件输入流
        Workbook hssfWorkbook = new HSSFWorkbook(inputStream);//获取工作簿
        Sheet sheet = hssfWorkbook.getSheetAt(0);//获取第一张表
        Row rowTitle = sheet.getRow(0);//获取标题行
        if (rowTitle != null) {
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int i = 0; i < cellCount; i++) {
                Cell cell = rowTitle.getCell(i);
                if(cell!=null){
                    String cellValue=cell.getStringCellValue();
                    System.out.print(cellValue);
                }
            }

        }
        int rowCount=sheet.getPhysicalNumberOfRows();
        for(int i=1;i<=rowCount;i++){
            Row rowData = sheet.getRow(i);
            if (rowData != null) {

                int cellCount=rowData.getPhysicalNumberOfCells();
                for(int j=0;j<=cellCount;j++){
                    Cell cell=rowData.getCell(i);
                    if(cell!=null){
                        CellType cellType =cell.getCellType();
                        String cellValue="";
                        switch (cellType){
                            case STRING:
                                System.out.println("[String]");
                                cellValue=cell.getStringCellValue();
                                break;
                            case NUMERIC:
                                System.out.println("numeric");
                                if(HSSFDateUtil.isCellDateFormatted(cell)){
                                    System.out.println("日期");
                                    Date date = cell.getDateCellValue();
                                    new DateTime(date).toString("yyyy-MM-dd");
                                }else{
                                    //不是日期格式，防止数字过长
                                    System.out.println("转换为字符串");
                                    cell.setCellType(CellType.STRING);
                                    cellValue=cell.toString();
                                }
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }

    }

    @Test
    public void evalTest() throws Exception{
        FileInputStream inputStream = new FileInputStream(PATH + "xx.xls");//文件输入流
        Workbook workbook = new HSSFWorkbook(inputStream);//获取工作簿
        Sheet sheet = workbook.getSheetAt(0);//获取第一张表
        Row row = sheet.getRow(4);//获取公式所在行
        Cell cell = row.getCell(0);//获取公式所在列
        FormulaEvaluator evaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);//获取公式
        CellType cellType = cell.getCellType();
        switch (cellType){
            case FORMULA:
                String formula = cell.getCellFormula();
                System.out.println(formula);
                CellValue cellValue = evaluator.evaluate(cell);
                System.out.println(cellValue);
                break;
        }

    }
}
