package com.atpl.service;

import java.util.Map;

public interface CsvUploadService {
	public <T> void importCsvToDatabase(String csvFilePath, Class<T> entityClass, Map<String, String> columnToFieldMapping,String sourceType);
//	public <T> void importExcelToDatabase(String excelFilePath, Class<T> entityClass,
//			Map<String, String> columnToFieldMapping, String sourceType);
}