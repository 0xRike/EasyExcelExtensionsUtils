package pers.rike.easyexcel.annotaion;

import java.lang.annotation.*;

/**
 * 将excel合并格合并成一行
 * @author rike
 */
@Documented
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelMergeRow {
  String parent() default "";
  String[] value() default {""};
}
