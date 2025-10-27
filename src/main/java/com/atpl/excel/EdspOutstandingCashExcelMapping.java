package com.atpl.excel;

import java.util.LinkedHashMap;
import java.util.Map;

public class EdspOutstandingCashExcelMapping implements ExcelMapping {

	  @Override
	    public Map<String, String> getColumnToFieldMapping() {
	        Map<String, String> mapping = new LinkedHashMap<>();

	        mapping.put("station_code", "stationCode");
	        mapping.put("tracking_id", "trackingId");
	        mapping.put("balance_due", "balanceDue");
	        mapping.put("cash_with_associate_dt", "cashWithAssociateDt"); 
	        mapping.put("age_bucket", "ageBucket");
	        mapping.put("status_code", "statusCode");
	        mapping.put("employee_name", "employeeName");
	        mapping.put("performed_by_2", "performedBy2");
	        mapping.put("channel", "channel");
	        mapping.put("type", "type");
	        mapping.put("zones", "zones");
	        mapping.put("order_id", "orderId");
	        mapping.put("payment_method", "paymentMethod");
	        mapping.put("billing_partner_code", "billingPartnerCode");

	        return mapping;
	    }
}