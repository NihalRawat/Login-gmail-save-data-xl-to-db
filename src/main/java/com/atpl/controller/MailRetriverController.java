package com.atpl.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.atpl.mail.service.EmailExcelProcessorService;

import lombok.RequiredArgsConstructor;

@RestController
@RequestMapping("/api/mail")
@CrossOrigin(origins = "*")
@RequiredArgsConstructor
public class MailRetriverController {
	@Autowired
	private EmailExcelProcessorService emailExcelProcessorService;

	@GetMapping("/read-email-excel")
	public String readExcelFromEmail() {
	    emailExcelProcessorService.downloadAndProcessAttachment();
	    return "Excel processed and data saved!";
	}
	
	@GetMapping("/folder-path-name")
	public String readFolderAndProcessExcelFiles(@RequestParam String folderPath) {
		try {
			String msg= emailExcelProcessorService.processAllExcelFiles(folderPath);
			return msg;
		}catch(Exception ex) {
			ex.printStackTrace();
			return "Error"+ex;
		}
	}
}
