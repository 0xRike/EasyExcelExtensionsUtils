package pers.rike.easyexcel.annotaion;

import java.lang.annotation.*;

/**
 * 被注释的属性会在导出时纵向合并
 * @author lixin
 */
@Documented
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCellMerge {

  String[] keywords() default {};

}
