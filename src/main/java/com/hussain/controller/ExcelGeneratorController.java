package com.hussain.controller;

import com.hussain.converter.ExcelMockData;
import com.hussain.response.ExcelVo;
import com.hussain.utils.ExcelUtil;
import lombok.AllArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.util.IOUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.List;
import java.util.concurrent.TimeUnit;

@Slf4j
@RestController
@AllArgsConstructor
@RequestMapping("/api")
public class ExcelGeneratorController {

    private final ExcelUtil excelUtil;

    private final ExcelMockData mockData;

    @GetMapping("/writeExcelManualFlush")
    public void writeExcelManualFlush(HttpServletResponse response) throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
        final long startTime = System.currentTimeMillis();
        List<ExcelVo> excelData = mockData.getExcelData();
        ByteArrayInputStream stream = excelUtil.writeExcelManualFlush(excelData);
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=manual_flush.xlsx");
        IOUtils.copy(stream, response.getOutputStream());
        final long endTime = System.currentTimeMillis();
        final long totalTime = endTime - startTime;
        final long hr = TimeUnit.MILLISECONDS.toHours(totalTime);
        final long min = TimeUnit.MILLISECONDS.toMinutes(totalTime)
                - TimeUnit.HOURS.toMinutes(TimeUnit.MILLISECONDS.toHours(totalTime));
        final long sec = TimeUnit.MILLISECONDS.toSeconds(totalTime)
                - TimeUnit.MINUTES.toSeconds(TimeUnit.MILLISECONDS.toMinutes(totalTime));
        final long ms = TimeUnit.MILLISECONDS.toMillis(totalTime)
                - TimeUnit.SECONDS.toMillis(TimeUnit.MILLISECONDS.toSeconds(totalTime));
        System.out.println(String.format(
                "Total time taken to execute 20000 records using manual flush: %d Hours %d Minutes %d Seconds %d Milliseconds",
                hr, min, sec, ms));
    }

    @GetMapping("/writeExcelAutoFlush")
    public void writeExcelAutoFlush(HttpServletResponse response) throws IOException, InvocationTargetException, IllegalAccessException, NoSuchMethodException {
        final long startTime = System.currentTimeMillis();
        List<ExcelVo> excelData = mockData.getExcelData();
        ByteArrayInputStream stream = excelUtil.writeExcelAutoFlush(excelData);
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=auto_flush.xlsx");
        IOUtils.copy(stream, response.getOutputStream());

        final long endTime = System.currentTimeMillis();
        final long totalTime = endTime - startTime;
        final long hr = TimeUnit.MILLISECONDS.toHours(totalTime);
        final long min = TimeUnit.MILLISECONDS.toMinutes(totalTime)
                - TimeUnit.HOURS.toMinutes(TimeUnit.MILLISECONDS.toHours(totalTime));
        final long sec = TimeUnit.MILLISECONDS.toSeconds(totalTime)
                - TimeUnit.MINUTES.toSeconds(TimeUnit.MILLISECONDS.toMinutes(totalTime));
        final long ms = TimeUnit.MILLISECONDS.toMillis(totalTime)
                - TimeUnit.SECONDS.toMillis(TimeUnit.MILLISECONDS.toSeconds(totalTime));

        System.out.println(String.format(
                "Total time taken to execute 20000 records using auto flush: %d Hours %d Minutes %d Seconds %d Milliseconds",
                hr, min, sec, ms));
    }
}
