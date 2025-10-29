package com.atpl.scheduled;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import com.atpl.mail.service.EmailExcelProcessorService;

import lombok.RequiredArgsConstructor;
@RequiredArgsConstructor
@Component
public class ScheduledTask {
	
	@Autowired
	private EmailExcelProcessorService emailExcelProcessorService;
	
	

//	@Scheduled(cron = "0 0 2 * * *")
	@Scheduled(cron = "*/30 * * * * *")
	public void scheduleFlipkartReport() {
		emailExcelProcessorService.downloadAndProcessAttachment();
	}
	
}
