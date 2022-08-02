package com.lvyou.micro.utils.excel;

import java.lang.annotation.*;

/**
 * @author kun.tan
 * @date 2021/1/6 19:36
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelExportField {

    String value() default "";

    int sort() default 0;

    int width() default 4700;

    String dataPattern() default "yyyy-MM-dd";
}
