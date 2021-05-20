package swagger2word.exporter;

import com.google.common.base.Joiner;
import com.google.common.base.Splitter;
import java.io.File;
import java.io.InputStream;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Optional;
import java.util.stream.Collectors;
import lombok.Setter;
import lombok.experimental.Accessors;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import swagger2word.App;
import swagger2word.config.Config;
import swagger2word.config.model.SheetConfig;
import swagger2word.exporter.model.excel.InterfaceInfo;
import swagger2word.exporter.model.excel.InterfaceStat;
import swagger2word.swagger.SwaggerReader;
import swagger2word.swagger.model.Endpoint;
import swagger2word.utils.Const;
import swagger2word.utils.Utils;
import swagger2word.writer.CellData;
import swagger2word.writer.Workbook;
import swagger2word.writer.Worksheet;

@Slf4j
@Accessors(chain = true)
public class InterfaceStatExporter implements Exporter {

  @Setter private int columnStartIndex;

  public InterfaceStatExporter getData(List<String> fileList) throws Exception {
    for (String file : fileList) {
    }

    sort();
    filter();
    update();

    return this;
  }

  private void update() throws Exception {

  }

  public void filter() {
    
  }

  private InterfaceStatExporter sort() {
    return this;
  }

  public InterfaceStatExporter write(String out) {
    try (InputStream is = App.class.getResourceAsStream("/templates/interface_stat.xlsx")) {
      Workbook wb = new Workbook();
      wb.loadTemplate(is);

      for (SheetConfig sheetConfig : Config.getExporterConfig(EXPORT_TYPE).getSheets()) {
        Worksheet ws = new Worksheet(wb);
        ws.setStart(2);
        ws.setSheet(wb.getWorkbook().getSheet(sheetConfig.getName()));

        List<String> serviceList = new ArrayList<>();
        for (String cloudAlias : sheetConfig.getClouds()) {
          serviceList.addAll(Config.getServiceListByCloud(cloudAlias));
        }

        List<InterfaceStat> serviceDataList =
            sortedDataList.stream()
                .filter(inf -> serviceList.contains(inf.getService()))
                .peek(
                    inf -> {
                      inf.setService(Config.getServiceName(inf.getService()));
                      inf.setIftype(Config.getIftypeName(inf.getIftype()));
                    })
                .collect(Collectors.toList());

        ws.fill(serviceDataList);
        ws.merge(serviceDataList, 0, 0, 0);
        ws.color(serviceDataList, 0);
        ws.highlight(serviceDataList, 3);

        // 处理公式
        for (Entry<String, String> formulaEntry : sheetConfig.getFormulas().entrySet()) {
          String outFieldName = formulaEntry.getKey();
          int outFieldIndex = wb.getRenderFields().indexOf(outFieldName);
          String expr = ws.parseExpr(formulaEntry.getValue(), formulaEntry.getKey());

          for (int i = 0; i < serviceDataList.size(); i++) {
            String eval = expr.replaceAll("@@", String.valueOf((i + ws.getStart() + 1)));
            log.debug("formula={}", eval);
            ws.formula(i + ws.getStart(), outFieldIndex, eval);
          }
        }
      }
      wb.write(String.valueOf(new File(out)));
    } catch (Exception e) {
      e.printStackTrace();
    }
    return this;
  }
}
