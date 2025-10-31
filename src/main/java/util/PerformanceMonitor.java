package util;

import java.time.Duration;
import java.time.Instant;
import java.util.concurrent.atomic.AtomicInteger;

public class PerformanceMonitor {

    private final Instant startTime;
    private final AtomicInteger filesProcessed;
    private final AtomicInteger studentsProcessed;

    public PerformanceMonitor() {
        this.startTime = Instant.now();
        this.filesProcessed = new AtomicInteger(0);
        this.studentsProcessed = new AtomicInteger(0);
    }

    public void incrementFilesProcessed() {
        filesProcessed.incrementAndGet();
    }

    public void addStudentsProcessed(int count) {
        studentsProcessed.addAndGet(count);
    }

    public void printSummary() {
        Duration duration = Duration.between(startTime, Instant.now());
        long seconds = duration.getSeconds();
        long millis = duration.toMillis() % 1000;

        System.out.println("\n" + "═".repeat(60));
        System.out.println("PERFORMANCE SUMMARY");
        System.out.println("═".repeat(60));
        System.out.println(String.format("Files Processed:    %d", filesProcessed.get()));
        System.out.println(String.format("Students Processed: %d", studentsProcessed.get()));
        System.out.println(String.format("Time Elapsed:       %d.%03d seconds", seconds, millis));
        System.out.println(String.format("Throughput:         %.2f files/sec",
                filesProcessed.get() / Math.max(1.0, seconds)));
        System.out.println("═".repeat(60));
    }
}
