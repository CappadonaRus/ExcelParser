package ru.excel;

public enum ReportType {

    FIRST("FIRST"),
    SECOND("SECOND");

    private final String name;

    ReportType(String name) {
        this.name = name;
    }

    @Override
    public String toString() {
        return name;
    }
}
