package pers.rike.easyexcel.entity;

import com.alibaba.excel.annotation.ExcelIgnoreUnannotated;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;
import pers.rike.easyexcel.annotaion.ExcelCellMerge;

@Data
@ExcelIgnoreUnannotated
public class AutoMergeEntity {

  /**
   * 业务板块序号
   */
  private Integer businessModuleNo;
  /**
   * 业务板块名称
   */
  @ExcelProperty(value = "业务板块", index = 0)
  @ExcelCellMerge
  private String businessModuleName;
  /**
   * 项目序号
   */
  private Integer projectNo;
  /**
   * 项目名称
   */
  @ExcelProperty(value = "项目", index = 1)
  @ExcelCellMerge
  private String projectName;
  /**
   * 子项目序号
   */
  private Integer childProjectNo;
  /**
   * 子项目名称
   */
  @ExcelProperty(value = "子项目", index = 2)
  @ExcelCellMerge
  private String childProjectName;
  /**
   * 子项目要求
   */
  @ExcelProperty(value = "要求", index = 3)
  @ExcelCellMerge
  private String childProjectDemand;
  /**
   * 子项目属性
   */
  @ExcelProperty(value = "属性", index = 4)
  @ExcelCellMerge
  private String childProjectAttribute;
  /**
   * 子项目反馈方式
   */
  @ExcelProperty(value = "反馈方式", index = 5)
  @ExcelCellMerge(keywords = {"单选", "多选"})
  private String childProjectFeedbackType;
  /**
   * 总分
   */
  @ExcelProperty(value = "总分", index = 6)
  @ExcelCellMerge(keywords = "")
  private String optionTotalPoints;
  /**
   * 得分
   */
  @ExcelProperty(value = "得分", index = 7)
  @ExcelCellMerge(keywords = "")
  private String optionScore;
  /**
   * 点检选项内容
   */
  @ExcelProperty(value = "选项", index = 8)
  private String checkOptionContent;
  /**
   * 点检选项结果
   */
  @ExcelProperty(value = "选择结果", index = 9)
  private String checkOptionResult;


}
