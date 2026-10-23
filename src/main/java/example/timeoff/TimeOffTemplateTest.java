package example.timeoff;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.Arrays;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.workday.aunit.AssemblyTestCase;
import com.workday.aunit.annotations.AssemblyTest;
import com.workday.aunit.annotations.AssertAfter;
import com.workday.aunit.annotations.AssertBefore;
import com.workday.aunit.annotations.UnitTest;

/** Run this class with Run As > AUnit Test in Workday Studio. */
@AssemblyTest(project = "Studio_ExcelTemplatePOC",
        displayLabel = "Populate a fictional Time Off XLSM and preserve its VBA")
public class TimeOffTemplateTest extends AssemblyTestCase {
    private static final String MIME = "application/vnd.ms-excel.sheet.macroEnabled.12";
    private static final String TEMPLATE = "WSAR-INF/time-off-template.xlsm";
    private byte[] originalTemplate;
    private boolean inputVerified;
    private boolean outputVerified;

    @UnitTest(startComponent = "Start")
    public void testPopulateTimeOffTemplate() throws Exception {
        inputVerified = false;
        outputVerified = false;
        try (InputStream input = getTestResourceInputStream(TEMPLATE)) {
            originalTemplate = readAll(input);
        }
        // This AUnit helper copies the resource's raw bytes into the message part.
        // Use a binary message part for XLSM; a Write/static-file text step changes it.
        setMessagePart(0, TEMPLATE, MIME);
    }

    @AssertBefore(id = "Populate", step = "PopulateTemplate")
    public void verifyBinaryTemplateInput() throws Exception {
        assertNotNull("The AUnit initialization method loaded the template", originalTemplate);
        byte[] actual = (byte[]) getMediationContext().getMessage().getMessagePart(0, byte[].class);
        assertTrue("Custom step receives exact source XLSM bytes", Arrays.equals(originalTemplate, actual));
        inputVerified = true;
        System.out.println("PASS: custom-step input equals the original template bytes.");
    }

    @AssertAfter(id = "Populate", step = "PopulateTemplate")
    public void verifyPopulatedWorkbookAndSave() throws Exception {
        assertTrue("Binary input assertion ran before the custom step", inputVerified);
        byte[] output = (byte[]) getMediationContext().getMessage().getMessagePart(0, byte[].class);
        assertFalse("The output contains populated data", Arrays.equals(originalTemplate, output));
        assertEquals("Macro-enabled output MIME type", MIME, getMediationContext().getMessage().getMimeType(0));

        String[][] rows = {
            {"000123", "Alex Sample", "2026-10-20", "PTO", "8"},
            {"000456", "Jordan Example", "2026-10-21", "SICK", "4"},
            {"000789", "Taylor Demo", "2026-10-22", "PTO", "8"}
        };
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(output))) {
            assertTrue("Workbook remains macro-enabled", workbook.isMacroEnabled());
            XSSFSheet sheet = workbook.getSheet("Time Off");
            assertNotNull("Time Off worksheet exists", sheet);
            for (int index = 0; index < rows.length; index++) {
                int row = index + 9;
                assertEquals("Employee ID in B" + (row + 1), rows[index][0], sheet.getRow(row).getCell(1).getStringCellValue());
                assertEquals("Name in C" + (row + 1), rows[index][1], sheet.getRow(row).getCell(2).getStringCellValue());
                double date = ChronoUnit.DAYS.between(LocalDate.of(1899, 12, 30), LocalDate.parse(rows[index][2]));
                assertEquals("Date in D" + (row + 1), date, sheet.getRow(row).getCell(3).getNumericCellValue(), 0.0);
                assertEquals("Type in E" + (row + 1), rows[index][3], sheet.getRow(row).getCell(4).getStringCellValue());
                assertEquals("Hours in F" + (row + 1), Double.parseDouble(rows[index][4]), sheet.getRow(row).getCell(5).getNumericCellValue(), 0.0);
            }
            assertEquals("Formula remains in F15", "SUM(F10:F12)", sheet.getRow(14).getCell(5).getCellFormula());
            assertEquals("Cached total in F15", 20.0, sheet.getRow(14).getCell(5).getNumericCellValue(), 0.0);
            assertTrue("VendorConfig worksheet remains hidden", workbook.isSheetHidden(workbook.getSheetIndex("VendorConfig")));
        }
        byte[] originalVba = zipEntry(originalTemplate, "xl/vbaProject.bin");
        byte[] outputVba = zipEntry(output, "xl/vbaProject.bin");
        assertTrue("VBA project bytes remain identical", Arrays.equals(originalVba, outputVba));

        // Studio's native AUnit policy permits its temporary directory. Avoid
        // reading custom system properties or scanning the source project.
        Path destination = Files.createTempFile("time-off-populated-", ".xlsm");
        Files.write(destination, output);
        assertTrue("Saved output equals mediated output bytes", Arrays.equals(output, Files.readAllBytes(destination)));
        outputVerified = true;
        System.out.println("PASS: all three rows, dates, 20 hours, formula, macro type, and VBA bytes verified.");
        System.out.println("VBA SHA-256: " + hex(MessageDigest.getInstance("SHA-256").digest(outputVba)));
        System.out.println("Populated workbook: " + destination.toAbsolutePath());
    }

    /** Allows the separate local validation harness to verify that AUnit ran assertions. */
    public boolean wasOutputVerified() { return outputVerified; }

    private static byte[] zipEntry(byte[] workbook, String name) throws IOException {
        try (ZipInputStream zip = new ZipInputStream(new ByteArrayInputStream(workbook))) {
            ZipEntry entry;
            while ((entry = zip.getNextEntry()) != null) {
                if (name.equals(entry.getName())) return readAll(zip);
            }
        }
        throw new IOException("Missing workbook part: " + name);
    }

    private static byte[] readAll(InputStream input) throws IOException {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        byte[] buffer = new byte[8192]; int count;
        while ((count = input.read(buffer)) != -1) output.write(buffer, 0, count);
        return output.toByteArray();
    }

    private static String hex(byte[] bytes) {
        StringBuilder result = new StringBuilder();
        for (byte value : bytes) result.append(String.format("%02x", value & 255));
        return result.toString();
    }
}
