package processor;

import model.StudentRecord;
import org.apache.poi.ss.usermodel.*;
import util.ExcelUtils;

import java.util.Map;
import java.util.Set;

public class SheetProcessor {

    public static int processSheet(Sheet sheet, String fileName,
                                   Set<String> processedEmails,
                                   Map<String, StudentRecord> studentMap,
                                   Map<String, String> nameToEmailMap,
                                   boolean pass1) {

        Row headerRow = findHeaderRow(sheet);
        if (headerRow == null) {
            return 0;
        }

        ColumnIndices indices = findColumnIndices(headerRow, sheet);

        if (pass1) {
            if (indices.emailIndex == -1) {
                return 0;
            }
        } else {
            if (indices.emailIndex != -1 || indices.nameIndex == -1) {
                return 0;
            }
        }

        return processRows(sheet, headerRow.getRowNum(), indices, processedEmails,
                studentMap, nameToEmailMap, pass1);
    }

    private static Row findHeaderRow(Sheet sheet) {
        for (int i = 0; i <= Math.min(20, sheet.getLastRowNum()); i++) {
            Row row = sheet.getRow(i);
            if (ExcelUtils.isLikelyHeaderRow(row)) {
                return row;
            }
        }
        return null;
    }

    private static ColumnIndices findColumnIndices(Row headerRow, Sheet sheet) {
        int nameIndex = -1;
        int emailIndex = -1;
        int possibleEmailIndex = -1;

        for (Cell cell : headerRow) {
            String headerValue = ExcelUtils.getCellValueAsString(cell).trim();
            if (headerValue.isEmpty()) continue;

            int colIndex = cell.getColumnIndex();

            if (nameIndex == -1 && ExcelUtils.isNameColumn(headerValue)) {
                nameIndex = colIndex;
            }

            if (emailIndex == -1 && ExcelUtils.isEmailColumn(headerValue)) {
                emailIndex = colIndex;
            }

            if (emailIndex == -1 && possibleEmailIndex == -1) {
                if (containsEmailInColumn(sheet, headerRow.getRowNum(), colIndex)) {
                    possibleEmailIndex = colIndex;
                }
            }
        }

        if (emailIndex == -1 && possibleEmailIndex != -1) {
            emailIndex = possibleEmailIndex;
        }

        return new ColumnIndices(nameIndex, emailIndex);
    }

    private static boolean containsEmailInColumn(Sheet sheet, int headerRowNum, int colIndex) {
        int emailCount = 0;
        int checkRows = Math.min(10, sheet.getLastRowNum() - headerRowNum);

        for (int i = 1; i <= checkRows; i++) {
            Row row = sheet.getRow(headerRowNum + i);
            if (row == null) continue;

            Cell cell = row.getCell(colIndex);
            String value = ExcelUtils.getCellValueAsString(cell).trim();

            if (ExcelUtils.isValidEmail(value)) {
                emailCount++;
            }
        }

        return emailCount >= 2;
    }

    private static int processRows(Sheet sheet, int startRow, ColumnIndices indices,
                                   Set<String> processedEmails,
                                   Map<String, StudentRecord> studentMap,
                                   Map<String, String> nameToEmailMap,
                                   boolean pass1) {
        int rowsProcessed = 0;

        for (int i = startRow + 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null || isEmptyRow(row)) continue;

            String name = extractName(row, indices.nameIndex);
            if (name.isEmpty()) continue;

            String normalizedName = normalizeName(name);
            String email;

            if (pass1) {
                email = extractEmail(row, indices.emailIndex);
                if (email.isEmpty() || !ExcelUtils.isValidEmail(email)) {
                    continue;
                }

                nameToEmailMap.putIfAbsent(normalizedName, email);
            } else {
                email = nameToEmailMap.get(normalizedName);
                if (email == null) {
                    email = ExcelUtils.generateEmailFromName(name);
                    nameToEmailMap.put(normalizedName, email);
                }
            }

            if (processedEmails.contains(email)) {
                continue;
            }

            processedEmails.add(email);
            updateStudentMap(studentMap, email, name);
            rowsProcessed++;
        }

        return rowsProcessed;
    }

    private static String normalizeName(String name) {
        return name.toLowerCase()
                .trim()
                .replaceAll("[^a-z]", "")
                .replaceAll("\\s+", "");
    }

    private static boolean isEmptyRow(Row row) {
        for (Cell cell : row) {
            String value = ExcelUtils.getCellValueAsString(cell).trim();
            if (!value.isEmpty() && !value.matches("\\d+")) {
                return false;
            }
        }
        return true;
    }

    private static String extractName(Row row, int nameIndex) {
        if (nameIndex == -1) return "";

        Cell nameCell = row.getCell(nameIndex);
        String name = ExcelUtils.getCellValueAsString(nameCell).trim();

        if (name.isEmpty() || name.matches("\\d+") || name.length() < 3) {
            return "";
        }

        return name;
    }

    private static String extractEmail(Row row, int emailIndex) {
        if (emailIndex == -1) return "";

        Cell emailCell = row.getCell(emailIndex);
        return ExcelUtils.getCellValueAsString(emailCell).trim().toLowerCase();
    }

    private static void updateStudentMap(Map<String, StudentRecord> studentMap,
                                         String email, String name) {
        studentMap.compute(email, (key, existing) -> {
            if (existing == null) {
                return new StudentRecord(name, email);
            } else {
                existing.incrementCount();
                if (!name.isEmpty() && !name.matches("\\d+")) {
                    existing.setName(name);
                }
                return existing;
            }
        });
    }

    private static class ColumnIndices {
        final int nameIndex;
        final int emailIndex;

        ColumnIndices(int nameIndex, int emailIndex) {
            this.nameIndex = nameIndex;
            this.emailIndex = emailIndex;
        }
    }
}
