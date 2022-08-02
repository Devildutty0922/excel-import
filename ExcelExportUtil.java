package com.lvyou.micro.utils.excel;

import cn.hutool.core.text.CharSequenceUtil;
import com.lvyou.micro.constant.ExcelConstants;
import com.lvyou.micro.exception.ApiException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.MediaType;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author kun.tan
 * @date 2021/1/6 19:34
 */
public class ExcelExportUtil {

    private ExcelExportUtil() {
    }

    public static <T> void exportExcel(String fileName, List<T> data, Class<T> clazz) {
        try (Workbook workbook = new XSSFWorkbook()) {
            createWorkbook(workbook, data, clazz);
            exportExcelFromWorkbook(workbook, fileName);
        } catch (Exception e) {
            throw new ApiException("导出失败");
        }
    }

    public static void exportExcelFromWorkbook(Workbook workbook, String fileName) {
        HttpServletResponse response = Optional.ofNullable(Optional.ofNullable(
                (ServletRequestAttributes) RequestContextHolder.getRequestAttributes())
                .orElseThrow(() -> new ApiException("获取response失败")).getResponse())
                .orElseThrow(() -> new ApiException("获取response失败"));
        try {
            response.setHeader("Content-Disposition", "attachment; filename=" +
                    URLEncoder.encode(fileName + ExcelConstants.ExcelSuffixName.NEW_VERSION, ExcelConstants.CharacterEncoding.CODE_UTF8));
            response.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
            response.setCharacterEncoding(ExcelConstants.CharacterEncoding.CODE_UTF8);
            ServletOutputStream outputStream = response.getOutputStream();
            workbook.write(outputStream);
            outputStream.close();
        } catch (IOException e) {
            throw new ApiException("导出失败");
        }
    }

    public static <T> void createWorkbook(Workbook workbook, List<T> data, Class<T> clazz) throws IllegalAccessException {
        // 创建工作表对象
        Sheet sheet = workbook.createSheet("sheet1");

        Field[] fields = clazz.getDeclaredFields();
        List<Field> fieldList = Arrays.stream(fields)
                .filter(field -> field.isAnnotationPresent(ExcelExportField.class))
                .sorted(Comparator.comparing(field -> field.getAnnotation(ExcelExportField.class).sort()))
                .collect(Collectors.toList());

        int[] columnWidthArray = new int[fieldList.size()];

        // 创建表头
        Row rowHeader = sheet.createRow(0);
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //设置预定义填充颜色
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.index);

        // 数据处理
        for (int i = 0; i < data.size(); i++) {
            //创建工作表的行(表头占用1行, 这里从第二行开始)
            Row row = sheet.createRow(i + 1);
            T t = data.get(i);
            // 填充列数据
            for (int j = 0; j < fieldList.size(); j++) {
                Field field = fieldList.get(j);
                field.setAccessible(Boolean.TRUE);
                Object value = fieldList.get(j).get(t);
                String dataValue = Optional.ofNullable(value).orElse("").toString();
                Class<?> type = field.getType();
                String dataPattern = field.getAnnotation(ExcelExportField.class).dataPattern();
                if (type == Date.class) {
                    dataValue = new SimpleDateFormat(dataPattern).format((Date) value);
                } else if (type == LocalDate.class || type == LocalDateTime.class) {
                    assert value != null;
                    dataValue = DateTimeFormatter.ofPattern(dataPattern).format((TemporalAccessor) value);
                }
                if (isNumber(type)) {
                    // 设置为数值类型
                    row.createCell(j).setCellValue(Double.parseDouble(dataValue));
                    continue;
                }
                row.createCell(j).setCellValue(dataValue);
                int tempWidth = CharSequenceUtil.bytes(dataValue, ExcelConstants.CharacterEncoding.GBK).length * 260;
                if (columnWidthArray[j] < tempWidth) {
                    columnWidthArray[j] = tempWidth;
                }
            }
        }

        // 表头处理
        for (int i = 0; i < fieldList.size(); i++) {
            Field field = fieldList.get(i);
            field.setAccessible(Boolean.TRUE);
            ExcelExportField annotation = field.getAnnotation(ExcelExportField.class);
            sheet.setColumnWidth(i, Math.min(Math.max(annotation.width(), columnWidthArray[i]), 30000));
            Cell cell = rowHeader.createCell(i);
            cell.setCellValue(annotation.value());
            cell.setCellStyle(cellStyle);
        }

        // 设置列格式
        for (int colIndex = 0; colIndex < fieldList.size(); colIndex++) {
            Field field = fieldList.get(colIndex);
            Class<?> type = field.getType();
            field.setAccessible(Boolean.TRUE);
            if (!isNumber(type)) {
                continue;
            }
            ExcelExportField annotation = field.getAnnotation(ExcelExportField.class);
            CellStyle numberStyle = workbook.createCellStyle();
            String fmt = isDefaultFormat(annotation.dataPattern()) ? getDefaultFormat(type) : annotation.dataPattern();
            numberStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat(fmt));
            for (int rowIndex = 1; rowIndex < data.size(); rowIndex++) {
                sheet.getRow(rowIndex).getCell(colIndex).setCellStyle(numberStyle);
            }
        }
    }


    private static boolean isDefaultFormat(String fmt) {
        return "yyyy-MM-dd HH:mm:ss".equals(fmt);
    }

    private static boolean isNumber(Class<?> type) {
        // 暂时只考虑这么多类型
        return Number.class.isAssignableFrom(type) || type == Integer.TYPE || type == Long.TYPE;
    }

    private static String getDefaultFormat(Class<?> type) {
        // 暂时只考虑这么多类型
        if (type == Integer.class || type == Integer.TYPE || type == Long.class || type == Long.TYPE) {
            return "0";
        }
        return "0.00";
    }
}
