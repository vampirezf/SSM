package com.ssm.test;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author zf
 * @date 2021/7/26 10:35
 */
public class ExcelMatchUtil {
    //手机号正则
    private static String REX_PHONE= "(1([358][0-9]|4[579]|66|7[0135678]|9[89])[0-9]{8})(?!(\\.))(?!(\\d+))";
    //身份证正则
    private static String REX_IDCARD= "([1-9]\\d{5}(18|19|20)\\d{2}((0[1-9])|(1[0-2]))(([0-2][1-9])|10|20|30|31)\\d{3}[0-9Xx])(?!(\\.))(?!(\\d{1,20}))";

    public static boolean ExcelToHtml(String file) {
        InputStream fis = null;
        try {
            //fis = new FileInputStream(file);
            fis =HttpFIleUtil.inputStream(file);
            Workbook workbook = null;
            if (file.endsWith(".xlsx")||file.endsWith(".XLSX")) {
                workbook = new XSSFWorkbook(fis);
            } else if (file.endsWith(".xls") ||file.endsWith(".XLS")|| file.endsWith(".et")||file.endsWith(".ET")) {
                workbook = new HSSFWorkbook(fis);
            }
            //fis.close();
            /* 读EXCEL文字内容 */
            List<String> matchPhone = new ArrayList<>();
            List<String> matchIdCard = new ArrayList<>();
            for (int i=0;i<workbook.getNumberOfSheets();i++){
                // 获取第一个sheet表，也可使用sheet表名获取
                Sheet sheet = workbook.getSheetAt(i);
                // 获取行
                Iterator<Row> rows = sheet.rowIterator();
                Row row;
                Cell cell;
                while (rows.hasNext()) {
                    row = rows.next();
                    // 获取单元格
                    Iterator<Cell> cells = row.cellIterator();
                    int num = 0;
                    while (cells.hasNext()) {
                        cell = cells.next();
                        num ++;
                        String cellValue = POIUtil.getCellValue(cell);
                        boolean phone = isPhone(cellValue);
                        boolean idCard = isIDCard(cellValue);
                        if(phone){
                            //matchPhone.add(cellValue);
                            return true;
                        }
                        if(idCard){
                            //matchIdCard.add(cellValue);
                            return true;
                        }
                    }
                }
            }
            return false;
            /*if(matchPhone.isEmpty() && !matchIdCard.isEmpty()){
                return "手机号和身份证号";
            }else if (!matchPhone.isEmpty() && matchIdCard.isEmpty()){
                return "手机号";
            }else if(matchPhone.isEmpty() && !matchIdCard.isEmpty()){
                return "身份证号";
            }else {
                return null;
            }*/
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (null != fis) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        //return null;
        return false;
    }

    /**
     * 验证手机号
     * @param str
     * @return
     */
    public static boolean isPhone(final String str) {
        Pattern p = null;
        Matcher m = null;
        boolean b = false;
        p = Pattern.compile(REX_PHONE);
        m = p.matcher(str);
        b = m.find();
        return b;
    }

    /**
     * 验证身份证
     * @param str
     * @return
     */
    public static boolean isIDCard(final String str) {
        Pattern p = null;
        Matcher m = null;
        boolean b = false;
        p = Pattern.compile(REX_IDCARD);
        m = p.matcher(str);
        b = m.find();
        return b;
    }
}
