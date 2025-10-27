package com.atpl.excel;

import java.util.LinkedHashMap;
import java.util.Map;

import com.atpl.excel.ExcelMapping;

public class EdspShortCashExcelMapping implements ExcelMapping {

    @Override
    public Map<String, String> getColumnToFieldMapping() {
        Map<String, String> mapping = new LinkedHashMap<>();

        mapping.put("bank_id", "bankId");
        mapping.put("station_code", "stationCode");
        mapping.put("station_name", "stationName");
        mapping.put("deposit_type", "depositType");
        mapping.put("bank_transaction_date", "bankTransactionDate");
        mapping.put("recon_date", "reconDate");
        mapping.put("pickup_date", "pickupDate");
        mapping.put("recon_file_name", "reconFileName");
        mapping.put("bank_amount", "bankAmount");
        mapping.put("deposit_expected_amount", "depositExpectedAmount");
        mapping.put("net_diff", "netDiff");
        mapping.put("sum_of_short_collection_amount", "sumOfShortCollectionAmount");
        mapping.put("sum_of_excess_collection_amount", "sumOfExcessCollectionAmount");
        mapping.put("manual_recon", "manualRecon");
        mapping.put("depositidamt", "depositIdAmt");
        mapping.put("type", "type");
        mapping.put("remarks", "remarks");
        mapping.put("noc_comments", "nocComments");
        mapping.put("bucket", "bucket");
        mapping.put("facility", "facility");
        mapping.put("zone", "zone");
        mapping.put("month", "month");
        mapping.put("noc_closure_remarks", "nocClosureRemarks");
        mapping.put("status", "status");
        mapping.put("sp_name", "spName");

        return mapping;
    }
}