package com.atpl.excel;

import java.util.LinkedHashMap;
import java.util.Map;


public class DspShortCashExcelMapping implements ExcelMapping {

    @Override
    public Map<String, String> getColumnToFieldMapping() {
        Map<String, String> mapping = new LinkedHashMap<>();

        mapping.put("remittance_id", "remittanceId");
        mapping.put("remittance_code", "remittanceCode");
        mapping.put("debrief_date", "debriefDate");
        mapping.put("station", "station");
        mapping.put("loc_type", "locType");
        mapping.put("loc_zone", "locZone");
        mapping.put("employee_id", "employeeId");
        mapping.put("employee_name", "employeeName");
        mapping.put("reason_code", "reasonCode");
        mapping.put("required_amount", "requiredAmount");
        mapping.put("submitted_short_excess", "submittedShortExcess");
        mapping.put("category", "category");
        mapping.put("week", "week");
        mapping.put("channel", "channel");
        mapping.put("PartnerCode", "partnerCode");
        mapping.put("load_dt", "loadDate");

        return mapping;
    }
}