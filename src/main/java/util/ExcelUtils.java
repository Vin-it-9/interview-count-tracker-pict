package util;

import org.apache.poi.ss.usermodel.*;

import java.util.Date;
import java.util.regex.Pattern;

public class ExcelUtils {

    private static final Pattern EMAIL_PATTERN = Pattern.compile(
            "^[A-Za-z0-9+_.-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}$"
    );

    public static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";

        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return formatDate(cell.getDateCellValue());
                    }
                    double numValue = cell.getNumericCellValue();
                    if (numValue == (long) numValue) {
                        return String.valueOf((long) numValue);
                    }
                    return String.valueOf(numValue);
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    return getFormulaValue(cell);
                case BLANK:
                    return "";
                default:
                    return "";
            }
        } catch (Exception e) {
            return "";
        }
    }

    private static String getFormulaValue(Cell cell) {
        try {
            return cell.getStringCellValue();
        } catch (IllegalStateException e) {
            try {
                double numValue = cell.getNumericCellValue();
                if (numValue == (long) numValue) {
                    return String.valueOf((long) numValue);
                }
                return String.valueOf(numValue);
            } catch (IllegalStateException e2) {
                return "";
            }
        }
    }

    private static String formatDate(Date date) {
        return date.toString();
    }

    public static boolean isNameColumn(String header) {
        if (header == null || header.isEmpty()) return false;

        String normalized = header.toLowerCase()
                .trim()
                .replaceAll("[^a-z0-9]", "");

        return normalized.equals("name") ||
                normalized.equals("fullname") ||
                normalized.equals("studentname") ||
                normalized.equals("candidatename") ||
                normalized.equals("applicantname") ||
                normalized.equals("employeename") ||
                normalized.equals("personname") ||
                normalized.contains("nameof") ||
                normalized.endsWith("name");
    }

    public static boolean isEmailColumn(String header) {
        if (header == null || header.isEmpty()) return false;

        String normalized = header.toLowerCase()
                .trim()
                .replaceAll("[^a-z0-9]", "");

        return normalized.equals("email") ||
                normalized.equals("emailaddress") ||
                normalized.equals("emailid") ||
                normalized.equals("mail") ||
                normalized.equals("mailid") ||
                normalized.equals("mailaddress") ||
                normalized.contains("email") ||
                normalized.contains("mail");
    }

    public static boolean isValidEmail(String email) {
        if (email == null || email.isEmpty()) return false;
        return EMAIL_PATTERN.matcher(email.trim()).matches();
    }

    public static String generateEmailFromName(String name) {
        if (name == null || name.trim().isEmpty()) {
            return "unknown" + System.currentTimeMillis() + "@generated.local";
        }

        String normalized = name.toLowerCase()
                .trim()
                .replaceAll("[^a-z0-9\\s]", "")
                .replaceAll("\\s+", ".");

        return normalized + "@generated.local";
    }

    public static boolean isLikelyHeaderRow(Row row) {
        if (row == null) return false;

        int nonEmptyCells = 0;
        int likelyHeaderCells = 0;

        for (Cell cell : row) {
            String value = getCellValueAsString(cell).trim();
            if (!value.isEmpty()) {
                nonEmptyCells++;

                if (isNameColumn(value) || isEmailColumn(value) ||
                        isCommonHeaderTerm(value)) {
                    likelyHeaderCells++;
                }
            }
        }

        return nonEmptyCells >= 1 && likelyHeaderCells >= 1;
    }

    private static boolean isCommonHeaderTerm(String value) {
        String normalized = value.toLowerCase().trim();
        return normalized.contains("id") ||
                normalized.contains("no") ||
                normalized.contains("number") ||
                normalized.contains("roll") ||
                normalized.contains("prn") ||
                normalized.contains("reg") ||
                normalized.contains("date") ||
                normalized.contains("status") ||
                normalized.contains("branch") ||
                normalized.contains("department") ||
                normalized.contains("class") ||
                normalized.contains("division") ||
                normalized.contains("time") ||
                normalized.contains("reporting");
    }
}
