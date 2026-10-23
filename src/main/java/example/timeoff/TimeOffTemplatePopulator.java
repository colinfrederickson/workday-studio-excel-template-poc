package example.timeoff;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorksheetDocument;

/**
 * Demonstrates updating a pre-existing XLSM template using Studio's POI 3.16.
 * The assembly supplies and receives raw workbook bytes. VBA is never executed.
 * This example targets only the included three-row fictional Time Off template.
 */
public final class TimeOffTemplatePopulator {
    private static final String SHEET = "Time Off";
    private static final int FIRST_ROW = 9;
    private static final int CAPACITY = 3;

    public static final class TimeOff {
        final String employeeId, employeeName, type;
        final LocalDate date;
        final double hours;
        public TimeOff(String id, String name, String date, String type, double hours) {
            this.employeeId = id;
            this.employeeName = name;
            this.date = LocalDate.parse(date);
            this.type = type;
            this.hours = hours;
        }
    }

    /** Opens a copy of a template and updates existing input cells only. */
    public byte[] populate(byte[] template, List<TimeOff> records) throws Exception {
        if (records.isEmpty() || records.size() > CAPACITY) {
            throw new IllegalArgumentException("Template accepts 1 to 3 records");
        }
        for (TimeOff record : records) {
            if (record.employeeId.isEmpty() || record.employeeName.isEmpty()
                    || !(record.type.equals("PTO") || record.type.equals("SICK"))
                    || !Double.isFinite(record.hours) || record.hours < 0 || record.hours > 24) {
                throw new IllegalArgumentException("Invalid fictional Time Off record");
            }
        }
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(template))) {
            Map<XSSFSheet, CTCols[]> originalColumns = columnDefinitions(workbook);
            if (!workbook.isMacroEnabled()) {
                throw new IllegalArgumentException("Expected a macro-enabled workbook");
            }
            XSSFSheet sheet = workbook.getSheet(SHEET);
            if (sheet == null) throw new IllegalArgumentException("Missing Time Off worksheet");
            for (int row = FIRST_ROW; row < FIRST_ROW + CAPACITY; row++) {
                for (int column = 1; column <= 5; column++) {
                    XSSFRow existingRow = sheet.getRow(row);
                    XSSFCell cell = existingRow == null ? null : existingRow.getCell(column);
                    if (cell == null || cell.getCellStyle().getIndex() == 0) {
                        throw new IllegalArgumentException("Missing preformatted input cell "
                                + new CellReference(row, column).formatAsString());
                    }
                    if (cell.getCellTypeEnum() == CellType.FORMULA) {
                        throw new IllegalArgumentException("Refusing to replace input formula");
                    }
                    cell.setCellType(CellType.BLANK);
                }
            }
            for (int i = 0; i < records.size(); i++) {
                TimeOff record = records.get(i);
                XSSFRow row = sheet.getRow(FIRST_ROW + i);
                row.getCell(1).setCellValue(record.employeeId);
                row.getCell(2).setCellValue(record.employeeName);
                // Native Excel date serial, computed without timezone conversion.
                long serial = java.time.temporal.ChronoUnit.DAYS.between(
                        LocalDate.of(1899, 12, 30), record.date);
                row.getCell(3).setCellValue((double) serial);
                row.getCell(4).setCellValue(record.type);
                row.getCell(5).setCellValue(record.hours);
            }
            // This evaluates ordinary worksheet formulas. It does not execute VBA.
            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
            // POI 3.16 changes explicit collapsed=false column flags while loading.
            // Restore the template's complete column definitions after processing.
            restoreColumnDefinitions(originalColumns);
            ByteArrayOutputStream output = new ByteArrayOutputStream();
            workbook.write(output);
            return output.toByteArray();
        }
    }

    private static Map<XSSFSheet, CTCols[]> columnDefinitions(XSSFWorkbook workbook) throws Exception {
        Map<XSSFSheet, CTCols[]> definitions = new LinkedHashMap<XSSFSheet, CTCols[]>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            XSSFSheet sheet = workbook.getSheetAt(i);
            try (InputStream originalXml = sheet.getPackagePart().getInputStream()) {
                definitions.put(sheet, WorksheetDocument.Factory.parse(originalXml).getWorksheet().getColsArray());
            }
        }
        return definitions;
    }

    private static void restoreColumnDefinitions(Map<XSSFSheet, CTCols[]> definitions) {
        for (Map.Entry<XSSFSheet, CTCols[]> entry : definitions.entrySet()) {
            entry.getKey().getCTWorksheet().setColsArray(entry.getValue());
        }
    }


    /** Entry point called by the Studio custom step in assembly.xml. */
    public byte[] process(byte[] template) throws Exception {
        return populate(template, Arrays.asList(
            new TimeOff("000123", "Alex Sample", "2026-10-20", "PTO", 8),
            new TimeOff("000456", "Jordan Example", "2026-10-21", "SICK", 4),
            new TimeOff("000789", "Taylor Demo", "2026-10-22", "PTO", 8)));
    }
}
