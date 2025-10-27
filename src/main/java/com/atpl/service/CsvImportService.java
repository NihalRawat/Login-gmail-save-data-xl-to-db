package com.atpl.service;

import java.util.Map;

public interface CsvImportService {
//	public <T> void importCsvToDatabase(String csvFilePath, Class<T> entityClass, Map<String, String> columnToFieldMapping);
	public <T> void importCsvToDatabase(String csvFilePath, Class<T> entityClass, Map<String, String> columnToFieldMapping,String sourceType);
}