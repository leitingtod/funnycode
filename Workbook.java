package swagger2word.writer;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.stream.Collectors;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Data
public class Workbook {

  private XSSFWorkbook workbook;

  private final List<Worksheet> worksheets = new ArrayList<>();

  private final List<String> renderFields = new ArrayList<>();
  private final List<CellStyle> renderFieldStyleList = new ArrayList<>();

  public Workbook() {
    workbook = new XSSFWorkbook();
  }

  public void loadTemplate(InputStream file) throws IOException {

    workbook = new XSSFWorkbook(file);

    XSSFSheet ws = workbook.getSheetAt(0);

    for (int i = 1; i < ws.getLastRowNum() + 1; i++) {
      XSSFRow row = ws.getRow(i);
      for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
        String cellValue = row.getCell(j).getStringCellValue();
        if (cellValue.contains("{") && cellValue.contains("}")) {
          renderFields.add(cellValue.replaceAll("[{.}]", ""));
          //          System.out.println(fields);
          renderFieldStyleList.add(row.getCell(j).getCellStyle());
        }
      }
    }
  }

  public void read(String file) {}

  public void write(String file) throws IOException {
    FileOutputStream out = new FileOutputStream(String.valueOf(Paths.get(file)));

    workbook.write(out);
    out.close();
  }

  public Worksheet newWorksheet(String name) {
    // Replace chars forbidden in worksheet names (backslahses and colons) by dashes
    String sheetName = name.replaceAll("[/\\\\?*\\]\\[:]", "-");

    // Maximum length worksheet name is 31 characters
    if (sheetName.length() > 31) {
      sheetName = sheetName.substring(0, 31);
    }

    synchronized (worksheets) {
      // If the worksheet name already exists, append a number
      int number = 1;
      Set<String> names = worksheets.stream().map(Worksheet::getName).collect(Collectors.toSet());
      while (names.contains(sheetName)) {
        String suffix = String.format(Locale.ROOT, "_%d", number);
        if (sheetName.length() + suffix.length() > 31) {
          sheetName = sheetName.substring(0, 31 - suffix.length()) + suffix;
        } else {
          sheetName += suffix;
        }
        ++number;
      }
      Worksheet worksheet = new Worksheet(this, sheetName);
      worksheets.add(worksheet);
      return worksheet;
    }
  }

  public CellStyle createCellStyle(CellStyle styleFrom) {
    CellStyle style = workbook.createCellStyle();
    style.cloneStyleFrom(styleFrom);
    return style;
  }
}
