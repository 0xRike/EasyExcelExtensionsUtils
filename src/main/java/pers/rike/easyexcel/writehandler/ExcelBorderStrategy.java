package pers.rike.easyexcel.writehandler;

import com.alibaba.excel.constant.OrderConstant;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.style.AbstractCellStyleStrategy;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.*;

/**
 * @author RiKe
 */
@Data
@NoArgsConstructor
public class ExcelBorderStrategy extends AbstractCellStyleStrategy {


  @Override
  protected void setHeadCellStyle(Cell cell, Head head, Integer relativeRowIndex) {
    CellStyle cellStyle = cell.getCellStyle();
    cellStyle.setBorderRight(BorderStyle.THIN);
    cellStyle.setBorderLeft(BorderStyle.THIN);
    cellStyle.setBorderBottom(BorderStyle.THIN);
    cellStyle.setBorderTop(BorderStyle.THIN);
    cell.setCellStyle(cellStyle);
  }


  @Override
  protected void setContentCellStyle(Cell cell, Head head, Integer relativeRowIndex) {
    CellStyle cellStyle = cell.getCellStyle();
    cellStyle.setBorderRight(BorderStyle.THIN);
    cellStyle.setBorderLeft(BorderStyle.THIN);
    cellStyle.setBorderBottom(BorderStyle.THIN);
    cellStyle.setBorderTop(BorderStyle.THIN);
    cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    cellStyle.setWrapText(true);
    cell.setCellStyle(cellStyle);
  }

  @Override
  public int order() {
    return OrderConstant.FILL_STYLE + 1;
  }
}
