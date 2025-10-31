package service;

import model.StudentRecord;
import processor.ExcelFileProcessor;
import util.PerformanceMonitor;

import java.io.File;
import java.util.*;
import java.util.concurrent.*;

public class InterviewTrackerService {
    private final int threadPoolSize;
    private final PerformanceMonitor monitor;
    private final Map<String, String> nameToEmailMap;

    public InterviewTrackerService(int threadPoolSize) {
        this.threadPoolSize = threadPoolSize;
        this.monitor = new PerformanceMonitor();
        this.nameToEmailMap = new ConcurrentHashMap<>();
    }

    public Map<String, StudentRecord> processFiles(File[] files) {
        Map<String, StudentRecord> studentMap = new ConcurrentHashMap<>();

        System.out.println("\n🔄 PHASE 1: Processing files with email columns...");
        processFilesWithEmails(files, studentMap);

        System.out.println("\n🔄 PHASE 2: Processing files without email columns...");
        processFilesWithoutEmails(files, studentMap);

        return studentMap;
    }

    private void processFilesWithEmails(File[] files, Map<String, StudentRecord> studentMap) {
        ExecutorService executor = Executors.newFixedThreadPool(threadPoolSize);

        try {
            List<CompletableFuture<ExcelFileProcessor.ProcessingResult>> futures = new ArrayList<>();

            for (File file : files) {
                CompletableFuture<ExcelFileProcessor.ProcessingResult> future =
                        CompletableFuture.supplyAsync(
                                () -> ExcelFileProcessor.processFilePass1(file, studentMap, nameToEmailMap),
                                executor);
                futures.add(future);
            }

            CompletableFuture.allOf(futures.toArray(new CompletableFuture[0])).join();

            printPhaseResults(futures, 1);

        } catch (Exception e) {
            System.err.println("Error during phase 1: " + e.getMessage());
        } finally {
            shutdownExecutor(executor);
        }
    }

    private void processFilesWithoutEmails(File[] files, Map<String, StudentRecord> studentMap) {
        ExecutorService executor = Executors.newFixedThreadPool(threadPoolSize);

        try {
            List<CompletableFuture<ExcelFileProcessor.ProcessingResult>> futures = new ArrayList<>();

            for (File file : files) {
                CompletableFuture<ExcelFileProcessor.ProcessingResult> future =
                        CompletableFuture.supplyAsync(
                                () -> ExcelFileProcessor.processFilePass2(file, studentMap, nameToEmailMap),
                                executor);
                futures.add(future);
            }

            CompletableFuture.allOf(futures.toArray(new CompletableFuture[0])).join();

            printPhaseResults(futures, 2);

        } catch (Exception e) {
            System.err.println("Error during phase 2: " + e.getMessage());
        } finally {
            shutdownExecutor(executor);
        }
    }

    private void printPhaseResults(List<CompletableFuture<ExcelFileProcessor.ProcessingResult>> futures, int phase) {
        System.out.println("\n" + "─".repeat(70));
        System.out.println("PHASE " + phase + " RESULTS");
        System.out.println("─".repeat(70));

        for (CompletableFuture<ExcelFileProcessor.ProcessingResult> future : futures) {
            try {
                ExcelFileProcessor.ProcessingResult result = future.get();
                if (result.getRowsProcessed() > 0) {
                    result.printSummary();
                    monitor.incrementFilesProcessed();
                    monitor.addStudentsProcessed(result.getRowsProcessed());
                }
            } catch (Exception e) {
                System.err.println("Error getting result: " + e.getMessage());
            }
        }
    }

    private void shutdownExecutor(ExecutorService executor) {
        executor.shutdown();
        try {
            if (!executor.awaitTermination(60, TimeUnit.SECONDS)) {
                executor.shutdownNow();
            }
        } catch (InterruptedException e) {
            executor.shutdownNow();
            Thread.currentThread().interrupt();
        }
    }

    public void printPerformanceMetrics() {
        monitor.printSummary();
        System.out.println("Name-to-Email mappings created: " + nameToEmailMap.size());
    }
}
