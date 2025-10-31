package model;

import java.util.concurrent.atomic.AtomicInteger;

public class StudentRecord {

    private volatile String name;
    private final String email;
    private final AtomicInteger totalCount;

    public StudentRecord(String name, String email) {
        this.name = name;
        this.email = email;
        this.totalCount = new AtomicInteger(1);
    }

    public void incrementCount() {
        this.totalCount.incrementAndGet();
    }

    public String getName() {
        return name;
    }

    public String getEmail() {
        return email;
    }

    public int getTotalCount() {
        return totalCount.get();
    }

    public void setName(String name) {
        if (name != null && !name.isEmpty() && !name.equals(this.email.split("@")[0])) {
            this.name = name;
        }
    }

    @Override
    public String toString() {
        return String.format("%s,%s,%d", name, email, totalCount.get());
    }
}
