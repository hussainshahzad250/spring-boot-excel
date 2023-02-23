package com.hussain.utils;

import com.hussain.response.ExcelVo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

@Component
public class ExcelUtil {

    public static String capitalizeInitialLetter(String s) {
        if (s.length() == 0)
            return s;
        return s.substring(0, 1).toUpperCase() + s.substring(1);
    }

    public <T> ByteArrayInputStream writeToExcel(List<T> data) {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream(); XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            List<String> fieldNames = getFieldNamesForClass(data.get(0).getClass());
            int rowCount = 0;
            int columnCount = 0;
            Row row = sheet.createRow(rowCount++);
            for (String fieldName : fieldNames) {
                Cell cell = row.createCell(columnCount++);
                cell.setCellValue(fieldName);
            }
            Class<? extends Object> classz = data.get(0).getClass();
            for (T t : data) {
                row = sheet.createRow(rowCount++);
                columnCount = 0;
                for (String fieldName : fieldNames) {
                    Cell cell = row.createCell(columnCount);
                    Method method = null;
                    try {
                        method = classz.getMethod("get" + capitalize(fieldName));
                    } catch (NoSuchMethodException nme) {
                        method = classz.getMethod("get" + fieldName);
                    }
                    Object value = method.invoke(t, (Object[]) null);
                    if (value != null) {
                        if (value instanceof String) {
                            cell.setCellValue((String) value);
                        } else if (value instanceof Long) {
                            cell.setCellValue((Long) value);
                        } else if (value instanceof Integer) {
                            cell.setCellValue((Integer) value);
                        } else if (value instanceof Double) {
                            cell.setCellValue((Double) value);
                        }
                    }
                    columnCount++;
                }
            }
            workbook.write(outputStream);
            return new ByteArrayInputStream(outputStream.toByteArray());
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private static List<String> getFieldNamesForClass(Class<?> clazz) {
        List<String> fieldNames = new ArrayList<String>();
        Field[] fields = clazz.getDeclaredFields();
        for (int i = 0; i < fields.length; i++) {
            fieldNames.add(fields[i].getName());
        }
        return fieldNames.stream().map(o -> o.substring(0, 1).toUpperCase() + o.substring(1)).collect(Collectors.toList());
    }

    private static String capitalize(String s) {
        if (s.length() == 0)
            return s;
        return s.substring(0, 1).toUpperCase() + s.substring(1);
    }

    public ByteArrayInputStream writeExcelManualFlush(List<ExcelVo> excelData) throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream(); SXSSFWorkbook wb = new SXSSFWorkbook(-1)) {
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
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream(); SXSSFWorkbook wb = new SXSSFWorkbook(SXSSFWorkbook.DEFAULT_WINDOW_SIZE)) {
            // keep 100 rows in memory, exceeding rows will be flushed to disk
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
        }
    }

}
