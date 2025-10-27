package com.atpl.model;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.util.*;

//public class CodPivotExcelGenerator {
//    public static void main(String[] args) throws Exception {
//        // Sample data
//        List<DataRow> rows = Arrays.asList(
//            new DataRow("CABTDayalnagarTempODH_DAY", "09-Apr", 43146, 43146),
//            new DataRow("CABTDayalnagarTempODH_DAY", "14-Apr", 62022, 0),
//            new DataRow("CABTDayalnagarTempODH_DAY", "15-Apr", 72034, 0),
//            new DataRow("CABTDayalnagarTempODH_DAY", "16-Apr", 48145, 0),
//            new DataRow("CABTDayalnagarTempODH_DAY", "17-Apr", 54674, 0),
//            new DataRow("CABTDayalnagarTempODH_DAY", "18-Apr", 54156, 0),
//            new DataRow("CABTDayalnagarTempODH_DAY", "19-Apr", 53265, 0),
//            new DataRow("CABTDayalnagarTempODH_DAY", "20-Apr", 57310, 0),
//            new DataRow("CABTDayalnagarTempODH_DAY", "21-Apr", 49946, 0),
//            new DataRow("CABTDayalnagarTempODH_DAY", "22-Apr", 59419, 0),
//            new DataRow("DependoAdaviSrirampurODH_ADR", "17-Apr", 43641, 0),
//            new DataRow("DependoAdaviSrirampurODH_ADR", "18-Apr", 61245, 0),
//            new DataRow("DependoAdaviSrirampurODH_ADR", "19-Apr", 79741, 0),
//            new DataRow("DependoAdaviSrirampurODH_ADR", "20-Apr", 76909, 0),
//            new DataRow("DependoAdaviSrirampurODH_ADR", "21-Apr", 69061, 0),
//            new DataRow("DependoAdaviSrirampurODH_ADR", "22-Apr", 95073, 0)
//        );
//
//        Map<String, Long> codPendingMap = new LinkedHashMap<>();
//        codPendingMap.put("CABTDayalnagarTempODH_DAY", 554117L);
//        codPendingMap.put("DependoAdaviSrirampurODH_ADR", 425670L);
//
//        Workbook workbook = new XSSFWorkbook();
//        Sheet sheet = workbook.createSheet("COD Pivot");
//
//        int rowIndex = 0;
//
//        // Header
//        Row header = sheet.createRow(rowIndex++);
//        header.createCell(0).setCellValue("South");
//        header.createCell(1).setCellValue("COD pendency as per InstaKart");
//        header.createCell(2).setCellValue("COD deposited reported on Dependo COD portal");
//
//        // Group rows by station
//        Map<String, List<DataRow>> grouped = new LinkedHashMap<>();
//        for (DataRow r : rows) {
//            grouped.computeIfAbsent(r.station, k -> new ArrayList<>()).add(r);
//        }
//
//        for (Map.Entry<String, List<DataRow>> entry : grouped.entrySet()) {
//            String station = entry.getKey();
//            List<DataRow> records = entry.getValue();
//
//            // Merge row for station
//            Row groupHeader = sheet.createRow(rowIndex);
//            groupHeader.createCell(0).setCellValue(station);
//            groupHeader.createCell(1).setCellValue(codPendingMap.getOrDefault(station, 0L));
//            groupHeader.createCell(2).setCellValue(records.get(0).portal); // Show 1st deposit if needed
//
//            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 2));
//
//            rowIndex++;
//
//            // Data rows
//            for (DataRow r : records) {
//                Row row = sheet.createRow(rowIndex++);
//                row.createCell(0).setCellValue(r.date);
//                row.createCell(1).setCellValue(r.client);
//                row.createCell(2).setCellValue(r.portal);
//            }
//        }
//
//        // Autosize columns
//        for (int i = 0; i < 3; i++) {
//            sheet.autoSizeColumn(i);
//        }
//
//        // Write to file
//        FileOutputStream fos = new FileOutputStream("C:\\Users\\Kunal\\Downloads\\Dependo Sample Excel\\COD_Pivot_Report.xlsx");
//        workbook.write(fos);
//        fos.close();
//        workbook.close();
//
//        System.out.println("Excel generated successfully.");
//    }
//
//    static class DataRow {
//        String station;
//        String date;
//        long client;
//        long portal;
//
//        public DataRow(String station, String date, long client, long portal) {
//            this.station = station;
//            this.date = date;
//            this.client = client;
//            this.portal = portal;
//        }
//    }
//}


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.*;

public class CodPivotExcelGenerator {

    public static void main(String[] args) {
        // Sample data
        Map<String, List<DailyCodData>> dataMap = new LinkedHashMap<>();

        dataMap.put("CABTDayalnagarTempODH_DAY", Arrays.asList(
                new DailyCodData("09-Apr", 43146, 43146),
                new DailyCodData("14-Apr", 62022, 0),
                new DailyCodData("15-Apr", 72034, 0),
                new DailyCodData("16-Apr", 48145, 0),
                new DailyCodData("17-Apr", 54674, 0),
                new DailyCodData("18-Apr", 54156, 0),
                new DailyCodData("19-Apr", 53265, 0),
                new DailyCodData("20-Apr", 57310, 0),
                new DailyCodData("21-Apr", 49946, 0),
                new DailyCodData("22-Apr", 59419, 0)
        ));

        dataMap.put("DependoAdaviSrirampurODH_ADR", Arrays.asList(
                new DailyCodData("17-Apr", 43641, 0),
                new DailyCodData("18-Apr", 61245, 0),
                new DailyCodData("19-Apr", 79741, 0),
                new DailyCodData("20-Apr", 76909, 0),
                new DailyCodData("21-Apr", 69061, 0),
                new DailyCodData("22-Apr", 95073, 0)
        ));

        // Output path
        String outputFilePath = "C:\\Users\\Kunal\\Downloads\\Dependo Sample Excel\\COD_Pivot_Report.xlsx";

        // Generate the report
        generateCodExcelReport(dataMap, outputFilePath);
    }

    static class DailyCodData {
        String date;
        long pendency;
        long deposited;

        public DailyCodData(String date, long pendency, long deposited) {
            this.date = date;
            this.pendency = pendency;
            this.deposited = deposited;
        }
    }

    public static void generateCodExcelReport(Map<String, List<DailyCodData>> dataMap, String outputFilePath) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("COD Report");

            // Header style for station names and total row
            CellStyle boldStyle = workbook.createCellStyle();
            Font boldFont = workbook.createFont();
            boldFont.setBold(true);
            boldStyle.setFont(boldFont);

            int rowIndex = 0;

            // Create column headers
            Row header = sheet.createRow(rowIndex++);
            header.createCell(0).setCellValue("South");
            header.createCell(1).setCellValue("COD pendency as per InstaKart");
            header.createCell(2).setCellValue("COD deposited reported on Dependo COD portal");

            // Iterate over each station group
            for (Map.Entry<String, List<DailyCodData>> entry : dataMap.entrySet()) {
                String stationName = entry.getKey();
                List<DailyCodData> entries = entry.getValue();

                // Calculate totals
                long totalPendency = entries.stream().mapToLong(d -> d.pendency).sum();
                long totalDeposited = entries.stream().mapToLong(d -> d.deposited).sum();

                // Station header row
                Row stationRow = sheet.createRow(rowIndex++);
                Cell stationCell = stationRow.createCell(0);
                stationCell.setCellValue(stationName);
                stationCell.setCellStyle(boldStyle);

                Cell pendencyTotalCell = stationRow.createCell(1);
                pendencyTotalCell.setCellValue(totalPendency);
                pendencyTotalCell.setCellStyle(boldStyle);

                Cell depositedTotalCell = stationRow.createCell(2);
                depositedTotalCell.setCellValue(totalDeposited);
                depositedTotalCell.setCellStyle(boldStyle);

                // Daily rows
                for (DailyCodData data : entries) {
                    Row row = sheet.createRow(rowIndex++);
                    row.createCell(0).setCellValue(data.date);
                    row.createCell(1).setCellValue(data.pendency);
                    row.createCell(2).setCellValue(data.deposited);
                }

                // Empty row between stations for readability
                rowIndex++;
            }

            // Autosize columns
            for (int i = 0; i < 3; i++) {
                sheet.autoSizeColumn(i);
            }

            // Write to file
            try (FileOutputStream out = new FileOutputStream(outputFilePath)) {
                workbook.write(out);
                System.out.println("Excel file created successfully at: " + outputFilePath);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
