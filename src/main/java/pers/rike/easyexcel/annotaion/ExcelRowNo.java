package pers.rike.easyexcel.annotaion;

import java.lang.annotation.*;

/**
 * 用于基础点检导入生成序号
 * @author lixin
 */
@Documented
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelRowNo {

  String target();

  String parent() default "";

}
