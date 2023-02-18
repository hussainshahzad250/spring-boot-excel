package com.hussain.utils;

import com.hussain.response.ExcelVo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;

@Component
public class ExcelUtil {

    public static String capitalizeInitialLetter(String s) {
        if (s.length() == 0)
            return s;
        return s.substring(0, 1).toUpperCase() + s.substring(1);
    }

    public ByteArrayInputStream writeExcelManualFlush(List<ExcelVo> excelData) throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
        SXSSFWorkbook wb = null;
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            wb = new SXSSFWorkbook(-1);
            Sheet sh = wb.createSheet();
            Class<ExcelVo> classz = (Class<ExcelVo>) excelData.get(0).getClass();
            Field[] fields = classz.getDeclaredFields();
            int noOfFields = fields.length;
            int rownum = 0;
            Row row = sh.createRow(rownum);
            for (int i = 0; i < noOfFields; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(fields[i].getName());
            }
            for (ExcelVo excelModel : excelData) {
                row = sh.createRow(rownum + 1);
                int colnum = 0;
                for (Field field : fields) {
                    String fieldName = field.getName();
                    Cell cell = row.createCell(colnum);
                    Method method = null;
                    try {
                        method = classz.getMethod("get" + capitalizeInitialLetter(fieldName));
                    } catch (NoSuchMethodException nme) {
                        method = classz.getMethod("get" + fieldName);
                    }
                    Object value = method.invoke(excelModel, (Object[]) null);
                    cell.setCellValue((String) value);
                    colnum++;
                }
                // manually control how rows are flushed to disk
                if (rownum % 100 == 0) {
                    // retain 100 last rows and flush all others
                    ((SXSSFSheet) sh).flushRows(100);
                }
                rownum++;
            }
            wb.write(outputStream);
            return new ByteArrayInputStream(outputStream.toByteArray());
        }
    }

    public ByteArrayInputStream writeExcelAutoFlush(List<ExcelVo> excelData) throws IOException, NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        SXSSFWorkbook wb = null;
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            // keep 100 rows in memory, exceeding rows will be flushed to disk
            wb = new SXSSFWorkbook(SXSSFWorkbook.DEFAULT_WINDOW_SIZE);
            Sheet sh = wb.createSheet();
            Class<ExcelVo> classZ = (Class<ExcelVo>) excelData.get(0).getClass();
            Field[] fields = classZ.getDeclaredFields();
            int noOfFields = fields.length;
            int rownum = 0;
            Row row = sh.createRow(rownum);
            for (int i = 0; i < noOfFields; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(fields[i].getName());
            }

            for (ExcelVo excelModel : excelData) {
                row = sh.createRow(rownum + 1);
                int colnum = 0;
                for (Field field : fields) {
                    String fieldName = field.getName();
                    Cell cell = row.createCell(colnum);
                    Method method = null;
                    try {
                        method = classZ.getMethod("get" + capitalizeInitialLetter(fieldName));
                    } catch (NoSuchMethodException nme) {
                        method = classZ.getMethod("get" + fieldName);
                    }
                    Object value = method.invoke(excelModel, (Object[]) null);
                    cell.setCellValue((String) value);
                    colnum++;
                }
                rownum++;
            }
            wb.write(outputStream);
            return new ByteArrayInputStream(outputStream.toByteArray());
        } finally {
            if (wb != null) {
                wb.close();
            }
        }
    }
}
