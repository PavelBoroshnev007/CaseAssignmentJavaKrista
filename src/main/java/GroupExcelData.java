import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.LinkedHashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GroupExcelData {
    public static void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream("input.xlsx");
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);

            Map<String, Double[]> groupedData = new LinkedHashMap<>();

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Skip header row
                Cell cellA = row.getCell(0);
                Cell cellB = row.getCell(1);

                if (cellA != null && cellB != null && cellA.getCellType() == CellType.NUMERIC && cellB.getCellType() == CellType.NUMERIC) {
                    String key = cellA.getNumericCellValue() + "-" + cellB.getNumericCellValue();

                    Double[] values = groupedData.get(key);
                    if (values == null) {
                        values = new Double[] {0.0, Double.MIN_VALUE};
                    }

                    values[0] += row.getCell(3).getNumericCellValue();

                    double currentValueE = row.getCell(4).getNumericCellValue();
                    if (currentValueE > values[1]) {
                        values[1] = currentValueE;
                    }

                    groupedData.put(key, values);
                }
            }

            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("Sheet1");

            Row headerRow = outputSheet.createRow(0);
            headerRow.createCell(0).setCellValue("A");
            headerRow.createCell(1).setCellValue("B");
            headerRow.createCell(2).setCellValue("D");
            headerRow.createCell(3).setCellValue("E");

            int rowNum = 1;
            for (Map.Entry<String, Double[]> entry : groupedData.entrySet()) {
                Row row = outputSheet.createRow(rowNum++);
                String[] keyParts = entry.getKey().split("-");
                Double[] values = entry.getValue();

                row.createCell(0).setCellValue(Double.parseDouble(keyParts[0]));
                row.createCell(1).setCellValue(Double.parseDouble(keyParts[1]));
                row.createCell(2).setCellValue(values[0]);
                row.createCell(3).setCellValue(values[1]);
            }

            FileOutputStream outputStream = new FileOutputStream("output.xlsx");
            outputWorkbook.write(outputStream);
            outputWorkbook.close();
            workbook.close();
            file.close();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}