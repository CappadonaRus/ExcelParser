package ru.excel.late_report.model;

public class CategoryRow {

    private String category;

    private String isEvaAnswered;

    private String country;

    public CategoryRow(String category, String isEvaAnswered, String country) {
        this.category = category;
        this.isEvaAnswered = isEvaAnswered;
        this.country = country;
    }

    public String getCategory() {
        return category;
    }

    public String getIsEvaAnswered() {
        return isEvaAnswered;
    }

    public String getCountry() {
        return country;
    }
}
