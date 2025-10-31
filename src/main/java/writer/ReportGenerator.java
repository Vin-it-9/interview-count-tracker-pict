package writer;

import model.StudentRecord;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ReportGenerator {

    public static void generateReport(Map<String, StudentRecord> studentMap, String outputPath) {
        if (studentMap.isEmpty()) {
            System.out.println("\n⚠️  No student records to export.");
            return;
        }

        List<StudentRecord> sortedStudents = new ArrayList<>(studentMap.values());
        sortedStudents.sort((s1, s2) -> Integer.compare(s2.getTotalCount(), s1.getTotalCount()));

        printConsoleReport(sortedStudents);
        exportToExcel(sortedStudents, outputPath);
    }

    private static void printConsoleReport(List<StudentRecord> students) {
        System.out.println("\n" + "═".repeat(60));
        System.out.println("FINAL REPORT (Sorted: Highest → Lowest Interview Count)");
        System.out.println("═".repeat(60));
        System.out.println(String.format("%-4s %-28s %-24s %s", "Rank", "Name", "Email", "Count"));
        System.out.println("─".repeat(60));

        int limit = Math.min(10, students.size());
        for (int i = 0; i < limit; i++) {
            StudentRecord student = students.get(i);
            System.out.println(String.format("%-4d %-28s %-24s %d",
                    (i + 1),
                    truncate(student.getName(), 28),
                    truncate(student.getEmail(), 24),
                    student.getTotalCount()));
        }

        if (students.size() > 10) {
            System.out.println("... and " + (students.size() - 10) + " more");
        }
        System.out.println("═".repeat(60));
    }

    private static void exportToExcel(List<StudentRecord> students, String outputPath) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Interview Report");

            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle dataStyle = createDataStyle(workbook);
            CellStyle rankStyle = createRankStyle(workbook);

            createHeaderRow(sheet, headerStyle);
            populateDataRows(sheet, students, dataStyle, rankStyle);
            autoSizeColumns(sheet);

            try (FileOutputStream fileOut = new FileOutputStream(outputPath)) {
                workbook.write(fileOut);
            }

            System.out.println("\n✓ Report exported to: " + outputPath);
            System.out.println("  Total unique students: " + students.size());

        } catch (IOException e) {
            System.err.println("❌ Error writing Excel report: " + e.getMessage());
        }
    }

    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);

        style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    private static CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    private static CellStyle createRankStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    private static void createHeaderRow(Sheet sheet, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(0);
        headerRow.setHeightInPoints(25);

        String[] headers = {"Rank", "Name", "Email", "Total_Count"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
    }

    private static void populateDataRows(Sheet sheet, List<StudentRecord> students,
                                         CellStyle dataStyle, CellStyle rankStyle) {
        for (int i = 0; i < students.size(); i++) {
            Row row = sheet.createRow(i + 1);
            StudentRecord student = students.get(i);

            Cell rankCell = row.createCell(0);
            rankCell.setCellValue(i + 1);
            rankCell.setCellStyle(rankStyle);

            Cell nameCell = row.createCell(1);
            nameCell.setCellValue(student.getName());
            nameCell.setCellStyle(dataStyle);

            Cell emailCell = row.createCell(2);
            emailCell.setCellValue(student.getEmail());
            emailCell.setCellStyle(dataStyle);

            Cell countCell = row.createCell(3);
            countCell.setCellValue(student.getTotalCount());
            countCell.setCellStyle(rankStyle);
        }
    }

    private static void autoSizeColumns(Sheet sheet) {
        for (int i = 0; i < 4; i++) {
            sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 1000);
        }
    }

    private static String truncate(String str, int maxLength) {
        return str.length() > maxLength ? str.substring(0, maxLength - 3) + "..." : str;
    }
}
