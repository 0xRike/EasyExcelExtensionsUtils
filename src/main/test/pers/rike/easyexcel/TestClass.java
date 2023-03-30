package pers.rike.easyexcel;


import cn.hutool.core.io.FileUtil;
import cn.hutool.json.JSONArray;
import cn.hutool.json.JSONUtil;
import cn.hutool.setting.Setting;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import org.junit.Test;
import pers.rike.easyexcel.entity.AutoMergeEntity;
import pers.rike.easyexcel.writehandler.ExcelAdaptiveWidthStrategy;
import pers.rike.easyexcel.writehandler.ExcelBorderStrategy;
import pers.rike.easyexcel.writehandler.ExcelFileCellMergeStrategy;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.List;

/**
 * 测试类
 */
public class TestClass {

  @Test
  public void ExcelCellMergeStrategyTest() throws URISyntaxException, IOException {
    File file = new File("src/main/resources/data/AutoMergeEntityData.json");
    JSONArray json = JSONUtil.readJSONArray(file, StandardCharsets.UTF_8);
    File outputPathName = new File("example", "autoMerge.xlsx");
    List<AutoMergeEntity> data = json.toList(AutoMergeEntity.class);
    try (ExcelWriter excel = EasyExcelFactory.write().head(AutoMergeEntity.class)
      .excelType(ExcelTypeEnum.XLSX)
      .inMemory(true)
      .registerWriteHandler(new ExcelAdaptiveWidthStrategy())
      .registerWriteHandler(new ExcelFileCellMergeStrategy())
      .registerWriteHandler(new ExcelBorderStrategy())
      .file(outputPathName).build()) {
      WriteSheet sheet = EasyExcelFactory.write().sheet().build();
      excel.write(data, sheet).finish();
    }
  }

}
