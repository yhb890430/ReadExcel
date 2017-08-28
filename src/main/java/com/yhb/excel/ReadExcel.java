package com.yhb.excel;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.net.URL;
import java.util.HashSet;
import java.util.Set;

public class ReadExcel {

    private static  final Logger LOGGER = Logger.getLogger(ReadExcel.class);

    private static final int STANDARDVALUE0 = 1000;
    private static final int STANDARDVALUE1 = 30000;
    private static final int STANDARDVALUE2 = 50000;
    private static final int STANDARDVALUE3 = 80000;

    static{
        //初始化LOG4J
        URL resource = ReadExcel.class.getResource("/log4j.properties");
        PropertyConfigurator.configure(resource);
    }

    public void readExcel(){

        try {
            FileInputStream fis = new FileInputStream(new File("D:/工装寿命管理.xlsx"));
            XSSFWorkbook workbook = new XSSFWorkbook(fis);

            XSSFSheet sheet = workbook.getSheetAt(0);

            Set<String> yellowSet = new HashSet<String>();
            Set<String> redSet = new HashSet<String>();

            for (Row row : sheet) {

                if(row.getRowNum() <3){
                    continue;
                }
                if(row == null){
                    continue;
                }
                Cell cell3 = row.getCell(3);
                Cell cell7 = row.getCell(7);

                Cell cell5 = row.getCell(5);
                Cell cell9 = row.getCell(9);

                if(cell3 != null){
                    CellType cellTypeEnum3 = cell3.getCellTypeEnum();
                    if(cellTypeEnum3 != null && cellTypeEnum3.equals(CellType.NUMERIC)){
                        if(cell5.getCellTypeEnum().equals(CellType.NUMERIC)){
                            Double value3 = cell3.getNumericCellValue();
                            Double value5 = cell5.getNumericCellValue();
                            orgData(value3,value5,yellowSet,redSet,row);
                        }

                    }
                }
                if(cell7 != null){
                    CellType cellTypeEnum7 = cell7.getCellTypeEnum();
                    if(cellTypeEnum7 != null && cellTypeEnum7.equals(CellType.NUMERIC)){
                        if(cell9.getCellTypeEnum().equals(CellType.NUMERIC)){
                            Double value7= cell7.getNumericCellValue();
                            Double value9 = cell9.getNumericCellValue();
                            orgData(value7,value9,yellowSet,redSet,row);
                        }
                    }
                }
            }

            //插入数据(工装更换，工装终止)
            int rowNum = sheet.getLastRowNum()+2;
            //合并单元格，用于填入一级标题
            CellRangeAddress cra=new CellRangeAddress(rowNum, rowNum, 0, 1);
            sheet.addMergedRegion(cra);
            XSSFRow row1 = sheet.createRow(rowNum);
            CellStyle yellowStyle =  createStyle(workbook,IndexedColors.YELLOW);
            XSSFCell cell = row1.createCell(0);
            cell.setCellValue("工装更换");
            cell.setCellStyle(yellowStyle);

            //二级标题
            rowNum += 1;
            XSSFRow row2 = sheet.createRow(rowNum);
            XSSFCell cell0 = row2.createCell(0);
            XSSFCell cell1 = row2.createCell(1);
            cell0.setCellValue("序号");
            cell1.setCellValue("物料号");

            int i = 1;
            createRow(yellowSet, rowNum ,i,sheet);

            //合并单元格，用于填入一级标题
            rowNum += 1;
            CellRangeAddress cra1=new CellRangeAddress(rowNum, rowNum, 0, 1);
            sheet.addMergedRegion(cra1);
            XSSFRow redRow = sheet.createRow(rowNum);
            CellStyle redStyle = createStyle(workbook,IndexedColors.RED);

            XSSFCell redRowCell = redRow.createCell(0);
            redRowCell.setCellValue("工装终止");
            redRowCell.setCellStyle(redStyle);
            XSSFRow sheetRow1 = sheet.createRow((rowNum += 1));
            XSSFCell cell3 = sheetRow1.createCell(0);
            XSSFCell cell4 = sheetRow1.createCell(1);
            cell3.setCellValue("序号");
            cell4.setCellValue("物料号");
            i = 1;
            createRow(redSet, rowNum ,i,sheet);
            saveExcel(workbook);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            LOGGER.error("未去读取到Excel文件！",e);
        } catch (IOException e) {
            LOGGER.error("读取Excel文件异常！",e);
            e.printStackTrace();
        }
    }

    private CellStyle createStyle(XSSFWorkbook workbook,IndexedColors color){
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(color.getIndex());
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private void createRow(Set<String> set,int rowNum,int i,XSSFSheet sheet){
        for(String value : set){
            rowNum += 1;
            XSSFRow row = sheet.createRow(rowNum);
            XSSFCell cell0 = row.createCell(0);
            XSSFCell cell1 = row.createCell(1);
            cell0.setCellValue(i);
            i++;
            cell1.setCellValue(value);
        }
    }

    private void orgData(Double value0,Double value1,Set<String> yellowSet,Set<String> redSet,Row row){
        switch (value0.intValue()){
            case 1000:
                if(value1.intValue() >= STANDARDVALUE0){
                    redSet.add(row.getCell(1).getStringCellValue());
                }else if(value1.intValue() >= 800 && value1.intValue() < STANDARDVALUE0){
                    yellowSet.add(row.getCell(1).getStringCellValue());
                }
                break;
            case 30000:
                if(value1.intValue() >= STANDARDVALUE1){
                    redSet.add(row.getCell(1).getStringCellValue());
                }else if(value1.intValue() >= 25000 && value1.intValue() < STANDARDVALUE1){
                    yellowSet.add(row.getCell(1).getStringCellValue());
                }
                break;
            case 50000:
                if(value1.intValue() >= STANDARDVALUE2){
                    redSet.add(row.getCell(1).getStringCellValue());
                }else if(value1.intValue() >= 45000 && value1.intValue() < STANDARDVALUE2){
                    yellowSet.add(row.getCell(1).getStringCellValue());
                }
                break;
            case 80000:
                if(value1.intValue() >= STANDARDVALUE3){
                    redSet.add(row.getCell(1).getStringCellValue());
                }else if(value1.intValue() >= 70000 && value1.intValue() < STANDARDVALUE3){
                    yellowSet.add(row.getCell(1).getStringCellValue());
                }
                break;
            default:
                break;
        }
    }

    /**
     * 保存工作薄
     * @param wb
     */
    private void saveExcel(XSSFWorkbook wb) {
        FileOutputStream fileOut;
        try {
            fileOut = new FileOutputStream("D:/工装寿命管理_new.xlsx");
            wb.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public void init(){

        JFrame frame = new JFrame("工装寿命计算");
        frame.setSize(new Dimension(400,400));
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        //创建面板，一个frame可以有多个面板，类似于div
        JPanel panel = new JPanel();
        // 添加面板
        frame.add(panel);


        //显示窗口
        frame.setVisible(true);
    }

    public static void main(String[] args) {

        ReadExcel readExcel = new ReadExcel();

        readExcel.readExcel();

    }

}
