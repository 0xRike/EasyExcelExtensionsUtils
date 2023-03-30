package pers.rike.easyexcel.writehandler;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.CellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.style.column.AbstractColumnWidthStyleStrategy;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * EasyExcel导出 自适应列宽策略
 *
 * @author lixin
 */
@Slf4j
public class ExcelAdaptiveWidthStrategy extends AbstractColumnWidthStyleStrategy {

  private Map<Integer, Map<Integer, Integer>> cache = new HashMap<>();
  private final int MAX_COL_WIDTH;
  private final int MIN_COL_WIDTH;

  public ExcelAdaptiveWidthStrategy(int maxColWidth, int minColWidth) {
    MAX_COL_WIDTH = maxColWidth;
    MIN_COL_WIDTH = minColWidth;
  }

  public ExcelAdaptiveWidthStrategy(int maxColWidth) {
    MAX_COL_WIDTH = maxColWidth;
    MIN_COL_WIDTH = 20;
  }

  public ExcelAdaptiveWidthStrategy() {
    this.MAX_COL_WIDTH = 80;
    this.MIN_COL_WIDTH = 5;
  }
  /**
   * 计算长度
   *
   * @param cellDataList
   * @param cell
   * @param isHead
   * @return
   */
  private Integer dataLength(List<WriteCellData<?>> cellDataList, Cell cell, Boolean isHead) {
    if (isHead) {
      return cell.getStringCellValue().getBytes().length;
    } else {
      CellData cellData = cellDataList.get(0);
      CellDataTypeEnum type = cellData.getType();
      if (type == null) {
        return -1;
      } else {
        switch (type) {
          case STRING:
            return Arrays.stream(cellData.getStringValue().split("\n")).mapToInt(e -> e.getBytes(StandardCharsets.UTF_16).length).max().orElse(MIN_COL_WIDTH);
          case BOOLEAN:
            return cellData.getBooleanValue().toString().getBytes(StandardCharsets.UTF_16).length;
          case NUMBER:
            return cellData.getNumberValue().toString().getBytes(StandardCharsets.UTF_16).length;
          default:
            return -1;
        }
      }
    }
  }

  /**
   * Sets the column width when head create
   */
  @Override
  protected void setColumnWidth(WriteSheetHolder writeSheetHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
    boolean needSetWidth = isHead || !CollUtil.isEmpty(cellDataList);
    if (needSetWidth) {
      Map<Integer, Integer> maxColumnWidthMap = cache.computeIfAbsent(writeSheetHolder.getSheetNo(), k -> new HashMap<>());

      Integer columnWidth = this.dataLength(cellDataList, cell, isHead)+3;
      if (columnWidth >= 0) {
        if (columnWidth > MAX_COL_WIDTH) {
          columnWidth = MAX_COL_WIDTH;
        }

        Integer maxColumnWidth = maxColumnWidthMap.get(cell.getColumnIndex());
        if (maxColumnWidth == null || columnWidth > maxColumnWidth) {
          maxColumnWidthMap.put(cell.getColumnIndex(), columnWidth);
          Sheet sheet = writeSheetHolder.getSheet();
          sheet.setColumnWidth(cell.getColumnIndex(), columnWidth * 256);
        }
      }
    }
  }
}