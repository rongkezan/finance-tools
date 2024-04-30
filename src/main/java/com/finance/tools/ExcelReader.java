package com.finance.tools;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.StrUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.Duration;
import java.time.LocalTime;
import java.util.*;

/**
 * 计算规则
 * 一天: 两两配对计算时差(精确到分钟)
 * 若打卡次数是奇数，则忽略最后一次打卡
 * 若打卡次数是2次，则减一小时
 * 一月: 将每日的分钟累加成最终的总时间
 */
public class ExcelReader {

    private static final BigDecimal HOUR = BigDecimal.valueOf(3600);

    public static void main(String[] args) {
        String excelFilePath = "D:\\workspace\\excel\\余姚市君超电器塑料厂_打卡时间表_20240301-20240331_计算v3.xlsx";
        try (FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
//                if (row.getRowNum() == 2) {
//                    for (Cell cell : row) {
//                        int columnIndex = cell.getColumnIndex();
//                        if (columnIndex >= 6) {
//                            System.out.print(cell.getStringCellValue() + "\t");
//                        }
//                    }
//                }
                if (row.getRowNum() < 3) {
                    continue;
                }
                System.out.print((row.getRowNum() + 1) + ":\t");
                BigDecimal totalSeconds = BigDecimal.ZERO;
                for (Cell cell : row) {
                    int columnIndex = cell.getColumnIndex();
                    String value = cell.getStringCellValue();
                    if (columnIndex == 0 || columnIndex == 1) {
                        System.out.print(value + "\t");
                    }
                    if (columnIndex >= 6) {
                        // 将单元格内的数据转换为时间格式
                        List<LocalTime> values = Arrays.stream(value.split("\n")).map(v -> {
                            if (StrUtil.isBlank(v)) {
                                return null;
                            }
                            v = v.replaceAll("外勤", "").trim();
                            return LocalTime.parse(v);
                        }).filter(Objects::nonNull).toList();
                        // 计算相差的秒数
                        BigDecimal seconds = diffMinute(values);
                        System.out.print(seconds.divide(HOUR, 1, RoundingMode.HALF_UP) + "\t");
                        totalSeconds = totalSeconds.add(seconds);
                    }
                }
                System.out.print(totalSeconds.divide(HOUR, 1, RoundingMode.HALF_UP) + "\t");
                System.out.print(totalSeconds.divide(HOUR, 1, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(15)) + "\t");
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static BigDecimal diffMinute(List<LocalTime> values) {
        if (CollUtil.isEmpty(values)) {
            return BigDecimal.ZERO;
        }
        LocalTime startTime = values.get(0);
        LocalTime endTime = values.get(values.size() - 1);
        Duration duration = Duration.between(startTime, endTime);
        long seconds = duration.getSeconds();
        seconds = seconds - 3600;
        // 开始和结束时间都在13点之前，开始和结束时间都在11点之后，则不需要减一小时
        if ((startTime.isBefore(LocalTime.of(13, 0)) && endTime.isBefore(LocalTime.of(13, 0)))
                || (startTime.isAfter(LocalTime.of(11, 0)) && endTime.isAfter(LocalTime.of(11, 0)))) {
            seconds = seconds + 3600;
        }
        long remainder = seconds % 1800;
        // 能整除则直接返回
        if (remainder == 0) {
            return BigDecimal.valueOf(seconds);
        }
        // 不能整除则判断余数是不是 > 20分钟，如果是，则向上取整，否则向下取整
        if (remainder > 1200) {
            seconds = (seconds / 1800) * 1800 + 1800;
        } else {
            seconds = (seconds / 1800) * 1800;
        }
        return BigDecimal.valueOf(seconds);
    }
}
