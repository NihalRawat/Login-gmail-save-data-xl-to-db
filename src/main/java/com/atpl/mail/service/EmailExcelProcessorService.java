package com.atpl.mail.service;

public interface EmailExcelProcessorService {

	void downloadAndProcessAttachment();
	 String processAllExcelFiles(String folderPath);
}
