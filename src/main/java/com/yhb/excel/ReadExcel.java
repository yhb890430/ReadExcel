package com.yhb.excel;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;

public class ReadExcel {

    private static  final Logger LOGGER = Logger.getLogger(ReadExcel.class);

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


            for (Row row : sheet) {

                if(row.getRowNum() <3){
                    continue;
                }

                for (Cell cell : row) {
                    int cloumnIndex = cell.getColumnIndex();
                    if(cloumnIndex == 5){
//                        System.out.println(cell.getNumericCellValue());
                        System.out.println(cell.getCellStyle().getFillBackgroundColor());
//                        System.out.println(cell.getCellStyle().getFillForegroundColor());
                    }
                }

            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
            LOGGER.error("未去读取到Excel文件！",e);
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    public static void main(String[] args) {

        ReadExcel readExcel = new ReadExcel();

        readExcel.readExcel();

    }

}
