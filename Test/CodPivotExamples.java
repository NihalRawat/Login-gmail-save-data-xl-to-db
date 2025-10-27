package com.atpl.model;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class CodPivotExamples {

    public static void main(String[] args) {
        // Sample data
        Map<String, StationData> dataMap = new LinkedHashMap<>();

        dataMap.put("CABTDayalnagarTempODH_DAY", new StationData(
                "South", "CABTDayalnagarTempODH_DAY", "ODH_DAY",
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
                "Central", "DependoAdaviSrirampurODH_ADR", "ODH_ADR",
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

            // Cell styles
            CellStyle boldStyle = workbook.createCellStyle();
            Font boldFont = workbook.createFont();
            boldFont.setBold(true);
            boldStyle.setFont(boldFont);

            int rowIndex = 0;

            // Create header row
            Row header = sheet.createRow(rowIndex);
            header.createCell(0).setCellValue("Zone");
            header.createCell(1).setCellValue("Station Name");
            header.createCell(2).setCellValue("Date");
            header.createCell(3).setCellValue("COD pendency as per InstaKart");
            header.createCell(4).setCellValue("COD deposited reported on Dependo COD portal");
            header.setRowStyle(boldStyle); // Apply bold style to the header row
            rowIndex++;

            for (StationData stationData : dataMap.values()) {
                int stationRowIndex = rowIndex; // Store the starting row index for the station
                Row stationRow = sheet.createRow(stationRowIndex);
                stationRow.createCell(0).setCellValue(stationData.zone);
                stationRow.createCell(1).setCellValue(stationData.stationName);
                rowIndex++; // Increment row index

                long totalPendency = 0;
                long totalDeposited = 0;
                int dataRowCount = 0; // Keep track of how many data rows we have for this station

                for (DailyCodData data : stationData.dailyData) {
                    Row dataRow = sheet.createRow(rowIndex++);
                    dataRow.createCell(2).setCellValue(data.date);
                    dataRow.createCell(3).setCellValue(data.pendency);
                    dataRow.createCell(4).setCellValue(data.deposited);
                    totalPendency += data.pendency;
                    totalDeposited += data.deposited;
                    dataRowCount++;
                }

                // Create a total row for each station
                Row totalRow = sheet.createRow(rowIndex++);
                totalRow.createCell(1).setCellValue("Total");  //  "Total" in  Station Name Column
                totalRow.createCell(3).setCellValue(totalPendency);
                totalRow.createCell(4).setCellValue(totalDeposited);

                totalRow.getCell(1).setCellStyle(boldStyle);
                totalRow.getCell(3).setCellStyle(boldStyle);
                totalRow.getCell(4).setCellStyle(boldStyle);

                // Merge cells for Station Name and Zone, but only if there's more than one row of data
                if (dataRowCount > 0) { // corrected condition.
                    sheet.addMergedRegion(new CellRangeAddress(stationRowIndex, rowIndex - 1, 0, 0));
                    sheet.addMergedRegion(new CellRangeAddress(stationRowIndex, rowIndex - 1, 1, 1));
                }
            }

            // Autosize columns
            for (int i = 0; i <= 4; i++) {
                sheet.autoSizeColumn(i);
            }

            // Write to the file
            try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
                workbook.write(fileOut);
                System.out.println("Excel file created successfully at: " + outputFilePath);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

