package com.atpl.utility;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

public class DateGenerator {
    public static List<String> getLastFiveDays() {
        List<String> dateHeaders = new ArrayList<>();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MMM-yyyy");

        // Generate last 5 days
        for (int i = 0; i < 5; i++) {
            dateHeaders.add(LocalDate.now().minusDays(i).format(formatter));
        }

        // Add ">5" column
        dateHeaders.add(">5");
        return dateHeaders;
    }

    public static void main(String[] args) {
        List<String> headers = getLastFiveDays();
        System.out.println(headers);
    }
}
