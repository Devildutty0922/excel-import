package com.lvyou.micro.utils;

import java.time.*;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.time.temporal.TemporalUnit;
import java.util.Date;

/**
 * LocalDateTime工具类
 *
 * @author kun.tan
 * @date 9:23 2021/11/23
 */
public class LocalDateTimeUtils {

    private static DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("[yyyy/MM/dd][yyyy-MM-dd]");
    private static DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("[yyyy-MM-dd HH:mm:ss][yyyy/MM/dd HH:mm:ss]");
    private static String defaultTimestr = "1900-01-01 00:00:00";

    private LocalDateTimeUtils() {
    }

    public static LocalDateTime getDefaultLocalDateTime() {
        return LocalDateTime.parse(defaultTimestr, dateTimeFormatter);
    }

    /**
     * Date转换为LocalDateTime
     *
     * @param date /
     * @return java.lang.String
     * @author kun.tan
     * @date 9:28 2021/11/23
     */
    public static LocalDateTime convertDateToLocalDateTime(Date date) {
        return LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault());
    }

    /**
     * LocalDate转换为LocalDateTime
     *
     * @param localDate /
     * @return java.lang.String
     * @author kun.tan
     * @date 9:28 2021/11/23
     */
    public static LocalDateTime convertLocalDateToLocalDateTime(LocalDate localDate) {
        if (null == localDate) {
            return null;
        }
        LocalTime time = LocalTime.of(0, 0, 0);
        return LocalDateTime.of(localDate, time);
    }

    /**
     * LocalDateTime转换为LocalDate
     *
     * @param localDateTime /
     * @return java.lang.String
     * @author kun.tan
     * @date 9:28 2021/11/23
     */
    public static LocalDate convertLocalDateTimeToLocalDate(LocalDateTime localDateTime) {
        if (null == localDateTime) {
            return null;
        }
        return localDateTime.toLocalDate();
    }

    /**
     * LocalDateTime转换为Date
     *
     * @param time /
     * @return java.lang.String
     * @author kun.tan
     * @date 9:28 2021/11/23
     */
    public static Date convertLocalDateTimeToDate(LocalDateTime time) {
        return Date.from(time.atZone(ZoneId.systemDefault()).toInstant());
    }
    /**
     * yyyy-MM-dd 转换为LocalDate
     *
     * @param str /
     * @return java.lang.String
     * @author kun.tan
     * @date 9:28 2021/11/23
     */
    public static LocalDate convertTimeStrToLocalDate(String str) {
        return LocalDate.parse(str, dateFormatter);
    }

    /**
     * yyyy-MM-dd HH:mm:ss转换为LocalDateTime
     *
     * @param str /
     * @return java.lang.String
     * @author kun.tan
     * @date 9:28 2021/11/23
     */
    public static LocalDateTime convertTimeStrToLocalDateTime(String str) {
        return LocalDateTime.parse(str, dateTimeFormatter);
    }

    /**
     * 判断是否为默认值时间
     *
     * @param time /
     * @return java.lang.String
     * @author kun.tan
     * @date 9:28 2021/11/23
     */
    public static Boolean isDefaultLocalDateTime(LocalDateTime time) {
        return time.equals(LocalDateTime.parse(defaultTimestr, dateTimeFormatter));
    }

    /**
     * 判断是否为默认值时间
     *
     * @param time /
     * @return java.lang.String
     * @author kun.tan
     * @date 9:28 2021/11/23
     */
    public static Boolean isDefaultLocalDate(LocalDate time) {
        LocalDateTime localDateTime = convertLocalDateToLocalDateTime(time);
        return LocalDateTime.parse(defaultTimestr, dateTimeFormatter).equals(localDateTime);
    }

    /**
     * 获取指定日期的毫秒
     *
     * @param time /
     * @return java.lang.String
     * @author kun.tan
     * @date 9:28 2021/11/23
     */
    public static Long getMilliByTime(LocalDateTime time) {
        return time.atZone(ZoneId.systemDefault()).toInstant().toEpochMilli();
    }

    /**
     * 获取指定日期的秒
     *
     * @param time /
     * @return java.lang.String
     * @author kun.tan
     * @date 9:28 2021/11/23
     */
    public static Long getSecondsByTime(LocalDateTime time) {
        return time.atZone(ZoneId.systemDefault()).toInstant().getEpochSecond();
    }

    /**
     * 获取指定时间的yyyy-MM-dd HH:mm:ss
     *
     * @param time /
     * @return java.lang.String
     * @author kun.tan
     * @date 9:28 2021/11/23
     */
    public static String formatTime(LocalDateTime time) {
        return time.format(dateTimeFormatter);
    }

    /**
     * 获取指定时间的指定格式
     *
     * @param pattern /
     * @return java.lang.String
     * @author kun.tan
     * @date 9:28 2021/11/23
     */
    public static String formatTime(LocalDateTime time, String pattern) {
        return time.format(DateTimeFormatter.ofPattern(pattern));
    }


    /**
     * 获取当前时间的指定格式
     *
     * @param pattern /
     * @return java.lang.String
     * @author kun.tan
     * @date 9:28 2021/11/23
     */
    public static String formatNow(String pattern) {
        return formatTime(LocalDateTime.now(), pattern);
    }

    /**
     * 日期加上一个数,根据field不同加不同值,field为ChronoUnit.*
     *
     * @param time, number, field
     * @return java.time.LocalDateTime
     * @author kun.tan
     * @date 9:27 2021/11/23
     */
    public static LocalDateTime plus(LocalDateTime time, long number, TemporalUnit field) {
        return time.plus(number, field);
    }

    /**
     * 日期减去一个数,根据field不同减不同值,field参数为ChronoUnit.*
     *
     * @param time, number, field
     * @return java.time.LocalDateTime
     * @author kun.tan
     * @date 9:26 2021/11/23
     */
    public static LocalDateTime minu(LocalDateTime time, long number, TemporalUnit field) {
        return time.minus(number, field);
    }

    /**
     * 获取两个日期的差  field参数为ChronoUnit.*
     *
     * @param startTime /
     * @param endTime   /
     * @param field     单位(年月日时分秒)
     * @return /
     */
    public static long betweenTwoTime(LocalDateTime startTime, LocalDateTime endTime, ChronoUnit field) {
        Period period = Period.between(LocalDate.from(startTime), LocalDate.from(endTime));
        if (field == ChronoUnit.YEARS) {
            return period.getYears();
        }
        if (field == ChronoUnit.MONTHS) {
            return period.getYears() * 12L + period.getMonths();
        }
        return field.between(startTime, endTime);
    }

    /**
     * 获取一天的开始时间，2017,7,22 00:00
     *
     * @param time /
     * @return /
     * @author kun.tan
     * @date 9:25 2021/11/23
     */
    public static LocalDateTime getDayStart(LocalDateTime time) {
        return time.withHour(0)
                .withMinute(0)
                .withSecond(0)
                .withNano(0);
    }

    /**
     * 获取一天的结束时间，2017,7,22 23:59:59.999999999
     *
     * @param time /
     * @return /
     * @author kun.tan
     * @date 9:25 2021/11/23
     */
    public static LocalDateTime getDayEnd(LocalDateTime time) {
        return time.withHour(23)
                .withMinute(59)
                .withSecond(59)
                .withNano(999999999);
    }
}