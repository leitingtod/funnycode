package swagger2word.writer;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import swagger2word.config.Config;
import swagger2word.exporter.model.excel.NumberFormat;
import swagger2word.utils.Utils;

@Data
@Slf4j
public class Worksheet {

  /** Maximum number of rows in Excel. */
  public static final int MAX_ROWS = 1_048_576;

  /** Maximum number of columns in Excel. */
  public static final int MAX_COLS = 16_384;

  private final Workbook workbook;
  private String name;
  private XSSFSheet sheet;
  private int start;
  private final List<XSSFRow> rows = new ArrayList<>();
  private final List<CellStyle> cellStyleList = new ArrayList<>();
  private List<String> mergeRegionList = new ArrayList<>();

  public Worksheet(Workbook workbook) {
    this.workbook = Objects.requireNonNull(workbook);

    for (CellStyle cellStyle : workbook.getRenderFieldStyleList()) {
      CellStyle style = workbook.createCellStyle(cellStyle);
      style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      cellStyleList.add(style);
    }
  }

  public Worksheet(Workbook workbook, String name) {
    this.workbook = Objects.requireNonNull(workbook);
    this.name = Objects.requireNonNull(name);
    sheet = workbook.getWorkbook().createSheet(name);
  }

  public void fill(List<?> objects) throws IllegalAccessException {
    if (objects == null || objects.size() < 1) {
      return;
    }

    int row = start;

    Class<?> clazz = objects.get(0).getClass();

    for (Object obj : objects) {
      int col = 0;
      for (String fieldName : workbook.getRenderFields()) {
        Field field = Utils.getObjectFields(clazz).get(fieldName);
        if (field == null) {
          log.error(
              "模板字段({}) 在类({})的属性集合{}中不存在，请检查模板是否填写正确",
              fieldName,
              clazz,
              Utils.getObjectFields(clazz).keySet());
          System.exit(1);
        }

        field.setAccessible(true);

        Object v = field.get(obj);

        CellStyle style = workbook.createCellStyle(workbook.getRenderFieldStyleList().get(col));

        NumberFormat numberFormat = field.getAnnotation(NumberFormat.class);

        log.debug("Cell[{}, {}]({})={}, format={}", row, col, field.getName(), v, numberFormat);

        if (numberFormat != null) {
          style.setDataFormat((short) BuiltinFormats.getBuiltinFormat(numberFormat.value()));
          value(row, col, Double.parseDouble(Optional.ofNullable(v).orElse("0.0").toString()));
        } else {
          if (v != null && Utils.isNumeric(v.toString())) {
            value(row, col, Double.parseDouble(v.toString()));
          } else {
            value(row, col, Optional.ofNullable(v).orElse("").toString());
          }
        }

        cell(row, col).setCellStyle(style);
        col++;
      }

      row++;
    }
  }

  public void merge(List<?> objects, int leftColIndex, int rightColIndex, int sentinelColIndex)
      throws IllegalAccessException {
    if (objects == null || objects.size() < 1) {
      return;
    }

    if (leftColIndex > rightColIndex) {
      rightColIndex = leftColIndex;
    }

    int begin = start;
    int end = start;

    Class<?> clazz = objects.get(0).getClass();

    String fieldName = workbook.getRenderFields().get(sentinelColIndex);
    Field field = Utils.getObjectFields(clazz).get(fieldName);
    field.setAccessible(true);

    Object lastV = field.get(objects.get(sentinelColIndex));

    int i = 0;
    for (Object obj : objects) {
      Object v = field.get(obj);
      i++;
      if (!v.equals(lastV)) {
        lastV = v;
        end = i;
        log.debug("merge cells({}, {}, {}, {})\n", begin, leftColIndex, end, leftColIndex);
        if (begin < end) {
          for (int j = leftColIndex; j <= rightColIndex; j++) {
            merge(range(begin, j, end, j));
            mergeRegionList.add(begin + "-" + end);
          }
        }
        begin = end + 1;
      }
    }
    end = objects.size() + 1;
    log.debug("merge cells({}, {}, {}, {})\n", begin, leftColIndex, end, leftColIndex);
    if (begin < end) {
      for (int j = leftColIndex; j <= rightColIndex; j++) {
        merge(range(begin, j, end, j));
        mergeRegionList.add(begin + "-" + end);
      }
    }
  }

  public void highlight(List<?> objects, int colIndex) throws IllegalAccessException {
    if (objects == null || objects.size() < 1) {
      return;
    }

    Class<?> clazz = objects.get(0).getClass();

    String fieldName = workbook.getRenderFields().get(colIndex);
    Field field = Utils.getObjectFields(clazz).get(fieldName);
    field.setAccessible(true);

    Map<String, CellStyle> cellStyleMap = new HashMap<>();

    for (String iftype : Config.getIftypeList()) {
      CellStyle style = workbook.createCellStyle(cellStyleList.get(colIndex));
      int[] rgb = Config.getIftypeConfigMap().get(iftype).getRGB();
      if (rgb.length == 3) {
        XSSFColor color = new XSSFColor(new java.awt.Color(rgb[0], rgb[1], rgb[2]));
        ((XSSFCellStyle) style).setFillForegroundColor(color);
        cellStyleMap.put(Config.getIftypeName(iftype), style);
      }
    }

    int i = start;
    for (Object obj : objects) {
      Object v = field.get(obj);
      if (v != null && v.toString() != null) {
        cell(i, colIndex).setCellStyle(cellStyleMap.get(v.toString()));
      }
      i++;
    }
  }

  public void color(List<?> objects, int sentinelColIndex) throws IllegalAccessException {
    if (objects == null || objects.size() < 1) {
      return;
    }

    Class<?> clazz = objects.get(0).getClass();

    String fieldName = workbook.getRenderFields().get(sentinelColIndex);
    Field field = Utils.getObjectFields(clazz).get(fieldName);
    field.setAccessible(true);

    Object lastV = field.get(objects.get(sentinelColIndex));

    int colNum = sheet.getRow(0).getPhysicalNumberOfCells();

    int colorCount = 1;

    int i = start;
    for (Object obj : objects) {
      Object v = field.get(obj);
      if (!v.equals(lastV)) {
        lastV = v;
        colorCount++;
      }
      if (colorCount % 2 == 0) {
        log.debug("color row({}): {}", i, cell(i, 1).getStringCellValue());
        for (int j = 0; j < colNum; j++) {
          cell(i, j).setCellStyle(cellStyleList.get(j));
        }
      }
      i++;
    }

    i = i - 1;
    if (colorCount % 2 == 0) {
      log.debug("color row({}): {}", i, cell(i, 1).getStringCellValue());
      for (int j = 0; j < colNum; j++) {
        cell(i, j).setCellStyle(cellStyleList.get(j));
      }
    }
  }

  public void formula(int row, int col, String expression) {
    cell(row, col).setCellFormula(expression);
  }

  public CellRangeAddress range(int top, int left, int bottom, int right) {
    // Check limits
    if (top < 0 || top >= Worksheet.MAX_ROWS || bottom < 0 || bottom >= Worksheet.MAX_ROWS) {
      throw new IllegalArgumentException();
    }
    if (left < 0 || left >= Worksheet.MAX_COLS || right < 0 || right >= Worksheet.MAX_COLS) {
      throw new IllegalArgumentException();
    }
    int _top = Math.min(top, bottom);
    int _left = Math.min(left, right);
    int _bottom = Math.max(bottom, top);
    int _right = Math.max(right, left);
    return new CellRangeAddress(_top, _bottom, _left, _right);
  }

  public void merge(CellRangeAddress address) {
    sheet.addMergedRegion(address);
  }

  private XSSFRow row(int row) {
    if (sheet.getRow(row) == null) {
      return sheet.createRow(row);
    }
    return sheet.getRow(row);
  }

  private XSSFCell cell(int row, int col) {
    if (row(row).getCell(col) == null) {
      return row(row).createCell(col);
    }
    return row(row).getCell(col);
  }

  public void value(int row, int col, boolean value) {
    cell(row, col).setCellValue(value);
  }

  public void value(int row, int col, double value) {
    cell(row, col).setCellValue(value);
  }

  public void value(int row, int col, Date value) {
    cell(row, col).setCellValue(value);
  }

  public void value(int row, int col, LocalDateTime value) {
    cell(row, col).setCellValue(value);
  }

  public void value(int row, int col, Calendar value) {
    cell(row, col).setCellValue(value);
  }

  public void value(int row, int col, String value) {
    row(row).createCell(col).setCellValue(value);
  }

  public void value(int row, int col, RichTextString value) {
    cell(row, col).setCellValue(value);
  }

  public void value(int row, int col, LocalDate value) {
    cell(row, col).setCellValue(value);
  }

  /**
   * Convert a column index to a column name.
   *
   * @param c Zero-based column index.
   * @return Column name.
   */
  public static String colToString(int c) {
    StringBuilder sb = new StringBuilder();
    while (c >= 0) {
      sb.append((char) ('A' + (c % 26)));
      c = (c / 26) - 1;
    }
    return sb.reverse().toString();
  }

  public static String getExprVar(String s) {
    if (s == null || (!s.contains("{.") && !s.contains("}"))) {
      return null;
    } else {
      return s.substring(s.indexOf("{") + 2, s.indexOf("}"));
    }
  }

  public static String getExprRange(String s) {
    if (s == null || (!s.contains("@<") && !s.contains(">@"))) {
      return null;
    } else {
      return s.substring(s.indexOf("@<") + 2, s.indexOf(">@"));
    }
  }

  public String parseExpr(String expr, String outFieldName) throws Exception {
    int outFieldIndex = workbook.getRenderFields().indexOf(outFieldName);
    String inFieldName = Worksheet.getExprVar(expr);
    int inFieldIndex = workbook.getRenderFields().indexOf(inFieldName);
    String rangeFieldName = Worksheet.getExprRange(expr);
    int rangeFieldIndex = workbook.getRenderFields().indexOf(rangeFieldName);

    if (outFieldIndex < 0) {
      throw new Exception(
          String.format("解析公式({%s}: {%s})输出变量错误，变量定义格式为 'outVar: {.InVar}'", outFieldName, expr));
    }

    if (inFieldName == null) {
      throw new Exception(
          String.format("解析公式({%s}: {%s})输入变量错误，变量定义格式为 'outVar: {.InVar}'", outFieldName, expr));
    }

    String parsed = expr.replace("{." + inFieldName + "}", colToString(inFieldIndex) + "@@");
    parsed =
        parsed.replace(
            "@<" + inFieldName + ">@",
            colToString(rangeFieldIndex) + "@<:" + colToString(rangeFieldIndex) + ">@");
    return parsed;
  }
}
