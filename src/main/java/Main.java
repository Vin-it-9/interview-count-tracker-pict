import model.StudentRecord;
import service.InterviewTrackerService;
import writer.ReportGenerator;

import java.io.File;
import java.util.Map;

public class Main {

    private static final String DEFAULT_FOLDER = "input_files";
    private static final String OUTPUT_FILE = "Interview_Appearance_Report.xlsx";
    private static final int THREAD_POOL_SIZE = Runtime.getRuntime().availableProcessors() * 2;

    public static void main(String[] args) {
        suppressLog4jWarnings();
        printBanner();

        String folderPath = args.length > 0 ? args[0] : DEFAULT_FOLDER;
        File[] files = getExcelFiles(folderPath);

        if (files == null || files.length == 0) {
            return;
        }

        InterviewTrackerService service = new InterviewTrackerService(THREAD_POOL_SIZE);
        Map<String, StudentRecord> studentMap = service.processFiles(files);

        ReportGenerator.generateReport(studentMap, OUTPUT_FILE);
        service.printPerformanceMetrics();
    }

    private static void suppressLog4jWarnings() {
        System.setProperty("log4j2.loggerContextFactory",
                "org.apache.logging.log4j.simple.SimpleLoggerContextFactory");
    }

    private static void printBanner() {
        System.out.println("\n" + "═".repeat(60));
        System.out.println("     CAMPUS INTERVIEW APPEARANCE TRACKER v2.0");
        System.out.println("═".repeat(60) + "\n");
    }

    private static File[] getExcelFiles(String folderPath) {
        File folder = new File(folderPath);

        if (!folder.exists() || !folder.isDirectory()) {
            System.err.println("❌ Invalid folder path: " + folderPath);
            System.out.println("\nUsage: java Main <folder_path>");
            return null;
        }

        File[] files = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".xlsx"));

        if (files == null || files.length == 0) {
            System.err.println("❌ No Excel (.xlsx) files found in: " + folderPath);
            return null;
        }

        System.out.println("📁 Folder: " + folderPath);
        System.out.println("📊 Found " + files.length + " Excel file(s)");
        System.out.println("🔄 Processing with " + THREAD_POOL_SIZE + " threads\n");
        System.out.println("─".repeat(60));

        return files;
    }
}
