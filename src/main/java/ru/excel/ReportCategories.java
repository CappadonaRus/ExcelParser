package ru.excel;

import java.util.ArrayList;
import java.util.List;

public class ReportCategories {

    private static List<String> categoriesList = new ArrayList<>();

    static {
        categoriesList.add("headers");
        for (int i = 1; i <= 125; i++) {
            categoriesList.add(String.valueOf(i));
        }
    }

    public static List<String> getCategoriesList() {
        return categoriesList;
    }
}
