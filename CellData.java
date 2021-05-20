package swagger2word.writer;

import java.time.LocalDateTime;
import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

@Data
public class CellData {

  private final Cell value;
  private String formula;
  private CellType type;
  private Object rawValue;

  public CellData(Cell cell) {
    value = cell;

    if (value == null) {
      return;
    }

    type = cell.getCellType();

    if (type.equals(CellType.NUMERIC)) {
      rawValue = cell.getNumericCellValue();
    } else if (type.equals(CellType.STRING)) {
      rawValue = cell.getStringCellValue();
    } else if (type.equals(CellType.FORMULA)) {
      formula = getFormula();
      if (cell.getCachedFormulaResultType().equals(CellType.NUMERIC)) {
        type = CellType.NUMERIC;
        rawValue = cell.getNumericCellValue();
      } else if (cell.getCachedFormulaResultType().equals(CellType.STRING)) {
        type = CellType.STRING;
        rawValue = cell.getStringCellValue();
      }
    } else if (type.equals(CellType.BOOLEAN)) {
      rawValue = cell.getBooleanCellValue();
    }
  }

  public Number asNumber() throws Exception {
    if (rawValue == null) {
      return 0.0;
    }
    requireType(CellType.NUMERIC);
    return (Number) rawValue;
  }

  public Boolean asBoolean() throws Exception {
    if (rawValue == null) {
      return false;
    }
    requireType(CellType.BOOLEAN);
    return (Boolean) rawValue;
  }

  public String asString() {
    if (rawValue == null) {
      return "";
    }
    // requireType(CellType.STRING);
    return rawValue.toString();
  }

  public LocalDateTime asDate() throws Exception {
    if (type == CellType.NUMERIC) {
      return null;
    } else if (type == CellType.BLANK) {
      return null;
    } else {
      throw new Exception("Wrong cell type " + type + " for date value");
    }
  }

  private void requireType(CellType requiredType) throws Exception {
    if (type != requiredType && type != CellType.BLANK) {
      throw new Exception("Wrong cell type " + type + ", wanted " + requiredType);
    }
  }
}
