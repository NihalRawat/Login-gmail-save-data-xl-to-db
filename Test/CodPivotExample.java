package com.atpl.model;

import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.util.*;
public class CodPivotExample {

    public static void main(String[] args) {
        // Sample data
        Map<String, StationData> dataMap = new LinkedHashMap<>();

        dataMap.put("CABTDayalnagarTempODH_DAY", new StationData(
                "South", "CABTDayalnagarTemp", "ODH_DAY",
                Arrays.asList(
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
                )));

        dataMap.put("DependoAdaviSrirampurODH_ADR", new StationData(
                "Central", "DependoAdaviSrirampur", "ODH_ADR",
                Arrays.asList(
                        new DailyCodData("17-Apr", 43641, 0),
                        new DailyCodData("18-Apr", 61245, 0),
                        new DailyCodData("19-Apr", 79741, 0),
                        new DailyCodData("20-Apr", 76909, 0),
                        new DailyCodData("21-Apr", 69061, 0),
                        new DailyCodData("22-Apr", 95073, 0)
                )));

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

    static class StationData {
        String zone;
        String stationName;
        String shName;
        List<DailyCodData> dailyData;

        public StationData(String zone, String stationName, String shName, List<DailyCodData> dailyData) {
            this.zone = zone;
            this.stationName = stationName;
            this.shName = shName;
            this.dailyData = dailyData;
        }
    }

    public static void generateCodExcelReport(Map<String, StationData> dataMap, String outputFilePath) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("COD Report");

            // Header style
            CellStyle boldStyle = workbook.createCellStyle();
            Font boldFont = workbook.createFont();
            boldFont.setBold(true);
            boldStyle.setFont(boldFont);

            int rowIndex = 0;

            // Column headers
            Row header = sheet.createRow(rowIndex++);
            header.createCell(0).setCellValue("Zone");
            header.createCell(1).setCellValue("Station Name");
            header.createCell(2).setCellValue("Short Name");
            header.createCell(3).setCellValue("Date");
            header.createCell(4).setCellValue("COD pendency as per InstaKart");
            header.createCell(5).setCellValue("COD deposited reported on Dependo COD portal");

            for (StationData stationData : dataMap.values()) {
                // Totals
                long totalPendency = stationData.dailyData.stream().mapToLong(d -> d.pendency).sum();
                long totalDeposited = stationData.dailyData.stream().mapToLong(d -> d.deposited).sum();

                // Header row for station
                Row stationRow = sheet.createRow(rowIndex++);
                stationRow.createCell(0).setCellValue(stationData.zone);
                stationRow.createCell(1).setCellValue(stationData.stationName);
                stationRow.createCell(2).setCellValue(stationData.shName);

                Cell pendencyCell = stationRow.createCell(4);
                pendencyCell.setCellValue(totalPendency);
                pendencyCell.setCellStyle(boldStyle);

                Cell depositedCell = stationRow.createCell(5);
                depositedCell.setCellValue(totalDeposited);
                depositedCell.setCellStyle(boldStyle);

                // Daily data rows
                for (DailyCodData data : stationData.dailyData) {
                    Row row = sheet.createRow(rowIndex++);
                    row.createCell(3).setCellValue(data.date);
                    row.createCell(4).setCellValue(data.pendency);
                    row.createCell(5).setCellValue(data.deposited);
                }

                rowIndex++; // Empty row
            }

            // Autosize
            for (int i = 0; i <= 5; i++) {
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
