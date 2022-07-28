package pers.rike.easyexcel.writehandler;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.util.ReflectUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.WorkbookWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import pers.rike.easyexcel.annotaion.ExcelCellMerge;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

/**
 * EasyExcel报表导出向上合并策略 <br/>
 * 请配合 @ExcelProperty 和 @ExcelIgnoreUnannotated 使用<br/>
 * @author rike
 */
public class ExcelFileCellMergeStrategy implements WorkbookWriteHandler {

  private final Map<Integer, Integer> mergeMap = new HashMap<>();
  private final Map<Integer, List<String>> keywordMap = new HashMap<>();

  /**
   * Called after all operations on the workbook have been completed
   *
   * @param writeWorkbookHolder
   */
  @Override
  public void afterWorkbookDispose(WriteWorkbookHolder writeWorkbookHolder) {
    writeWorkbookHolder.getHasBeenInitializedSheetIndexMap().forEach((k, v) -> {
      execMerge(v.getSheet(), v.getClazz(), v.getExcelWriteHeadProperty().getHeadMap());
    });
  }

  /**
   * 执行合并
   * @param sheet sheet
   * @param clazz 类信息
   * @param headMap EasyExcel 中的头部信息
   */
  @SuppressWarnings("rawtypes")
  public void execMerge(Sheet sheet, Class clazz, Map<Integer, Head> headMap) {
    //去除不合并的列
    Iterator<Map.Entry<Integer, Head>> iterator = headMap.entrySet().iterator();
    while (iterator.hasNext()) {
      Head head = iterator.next().getValue();
      Field field = ReflectUtil.getField(clazz, head.getFieldName());
      if (field.isAnnotationPresent(ExcelProperty.class) && field.isAnnotationPresent(ExcelCellMerge.class)) {
        // todo 将这里的 1 换成headRowNumber
//        mergeMap.put(head.getColumnIndex(), 1);
        keywordMap.put(head.getColumnIndex(), Arrays.stream(field.getDeclaredAnnotation(ExcelCellMerge.class).keywords()).filter(StrUtil::isNotEmpty).collect(Collectors.toList()));
      } else {
        iterator.remove();
      }
    }
    //合并
    try {
      List<Integer> colList = headMap.keySet().stream().sorted().collect(Collectors.toList());
      for (int colIndex : colList) {
        if (colIndex == colList.get(0)) {
          addMerge(sheet, 1, sheet.getLastRowNum(), colIndex);
        } else {
          List<CellRangeAddress> cellRangeAddresses = sheet.getMergedRegions().stream().filter(e -> e.getFirstColumn() == colIndex - 1).collect(Collectors.toList());
          for (CellRangeAddress e : cellRangeAddresses) {
            addMerge(sheet, e.getFirstRow(), e.getLastRow(), colIndex);
          }
        }
      }
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  /**
   * 判断是否可以合并 (由于该类为向上合并只与一个列相关)
   * @param sheet sheet
   * @param firstRow 首行
   * @param lastRow 尾行
   * @param curCol 列
   */
  public void addMerge(Sheet sheet, int firstRow, int lastRow, int curCol) {
    int tempIndex = firstRow;
    for (int i = firstRow; i < lastRow; i++) {
      Cell curCell = sheet.getRow(i).getCell(curCol);
      Cell downCell = sheet.getRow(i + 1).getCell(curCol);
      boolean equalsFlag = StrUtil.equals(downCell.getStringCellValue(), curCell.getStringCellValue());
      boolean keywordFlag = CollectionUtil.isEmpty(keywordMap.get(curCol)) || keywordMap.get(curCol).contains(curCell.getStringCellValue());
      boolean indexFlag = curCell.getRowIndex() != tempIndex;
      if (!equalsFlag || !keywordFlag) {
        if (indexFlag) {
          addRegion(sheet, tempIndex, curCell.getRowIndex(), curCol, curCol);
        }
        tempIndex = downCell.getRowIndex();
      }
      if (i == lastRow - 1) {
        addRegion(sheet, tempIndex, downCell.getRowIndex(), curCol, curCol);
      }
    }
  }

  /**
   * 将合并信息加入到sheet中
   * @param sheet sheet
   * @param firstRow 首行
   * @param lastRow 尾行
   * @param firstCol 首列
   * @param lastCol 尾列
   */
  public void addRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
    if (firstRow < lastRow && firstCol <= lastCol) {
      sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }
  }

}