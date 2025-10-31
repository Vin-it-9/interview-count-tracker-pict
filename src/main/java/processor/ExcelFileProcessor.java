package processor;

import model.StudentRecord;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

public class ExcelFileProcessor {

    public static ProcessingResult processFilePass1(File file,
                                                    Map<String, StudentRecord> globalStudentMap,
                                                    Map<String, String> nameToEmailMap) {
        ProcessingResult result = new ProcessingResult(file.getName());

        try {
            processFileInternal(file, globalStudentMap, nameToEmailMap, true, result);
        } catch (Exception e) {
            result.setError(e.getMessage());
        }

        return result;
    }

    public static ProcessingResult processFilePass2(File file,
                                                    Map<String, StudentRecord> globalStudentMap,
                                                    Map<String, String> nameToEmailMap) {
        ProcessingResult result = new ProcessingResult(file.getName());

        try {
            processFileInternal(file, globalStudentMap, nameToEmailMap, false, result);
        } catch (Exception e) {
            result.setError(e.getMessage());
        }

        return result;
    }

    private static void processFileInternal(File file,
                                            Map<String, StudentRecord> globalStudentMap,
                                            Map<String, String> nameToEmailMap,
                                            boolean pass1,
                                            ProcessingResult result) throws Exception {

        Set<String> processedInFile = ConcurrentHashMap.newKeySet();

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            int numberOfSheets = workbook.getNumberOfSheets();
            result.setTotalSheets(numberOfSheets);

            for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);

                if (sheet.getPhysicalNumberOfRows() == 0) {
                    continue;
                }

                int rowsProcessed = SheetProcessor.processSheet(
                        sheet, file.getName(), processedInFile, globalStudentMap,
                        nameToEmailMap, pass1);

                result.addProcessedRows(rowsProcessed);
                if (rowsProcessed > 0) {
                    result.incrementSheetsProcessed();
                }
            }

            result.setSuccess(true);
        }
    }

    public static class ProcessingResult {
        private final String fileName;
        private int totalSheets;
        private int sheetsProcessed;
        private int rowsProcessed;
        private boolean success;
        private String error;

        public ProcessingResult(String fileName) {
            this.fileName = fileName;
        }

        public void setTotalSheets(int totalSheets) {
            this.totalSheets = totalSheets;
        }

        public void incrementSheetsProcessed() {
            this.sheetsProcessed++;
        }

        public void addProcessedRows(int rows) {
            this.rowsProcessed += rows;
        }

        public void setSuccess(boolean success) {
            this.success = success;
        }

        public void setError(String error) {
            this.error = error;
            this.success = false;
        }

        public String getFileName() {
            return fileName;
        }

        public int getRowsProcessed() {
            return rowsProcessed;
        }

        public boolean isSuccess() {
            return success;
        }

        public void printSummary() {
            if (success && rowsProcessed > 0) {
                System.out.println(String.format("✓ %s: %d students from %d sheet(s)",
                        fileName, rowsProcessed, sheetsProcessed));
            }
        }
    }
}
