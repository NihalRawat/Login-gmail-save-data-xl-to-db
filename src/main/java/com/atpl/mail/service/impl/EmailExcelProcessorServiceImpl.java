package com.atpl.mail.service.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.mail.BodyPart;
import javax.mail.Folder;
import javax.mail.FolderClosedException;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Part;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import javax.transaction.Transactional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import com.atpl.entity.TblDsdpBillingTransaction;
import com.atpl.mail.service.EmailExcelProcessorService;
import com.atpl.repository.ExcelTransactionRepository;
import com.atpl.utility.HelperExtension;

@Service
public class EmailExcelProcessorServiceImpl implements EmailExcelProcessorService {

	@Value("${EXCEL_STORAGE_PATH}")
	private String excelPath;

	@Autowired
	private ExcelTransactionRepository excelTransactionRepository;

	@Value("${spring.mail.username}")
	private String username2;

	@Value("${spring.password}")
	private String password2;
	@Value("${POP_HOST_VALUE}")
	private String host2;

	@Value("${POP_PORT_VALUE}")
	private String port2;

	@Value("${POP_PROTOCOl}")
	private String protocol;


	@Value("${spring.password}")
	private String password;

	@Value("${spring.host}")
	private String host;

	@Value("${spring.port}")
	private String port;
	@Value("${email}")
	private String email;
	
	@Value("${spring.excel.sheet.name}")
	private String excelSheetName;
	
	@Value("${max.emails.to.check}")
	private int maxEmailsToChecks;
	
//	@Value("${folderPath}")
//	private String folderPath;

	@Override
	public void downloadAndProcessAttachment() {
		Properties properties = new Properties();
		properties.put("mail.store.protocol", protocol);
		 if (protocol.equalsIgnoreCase("imaps")) {
		        // Gmail IMAP configuration
		        properties.put("mail.imaps.host", host2);
		        properties.put("mail.imaps.port", port2);		
		        properties.put("mail.imaps.ssl.enable", "true");
		    } else {
		        // POP3 (old altruist server)
		        properties.put("mail.pop3s.host", host2);
		        properties.put("mail.pop3s.port", port2);
		        properties.put("mail.imaps.ssl.enable", "true");
		    }
		 
		boolean mailProcessed = false;
		Store store = null;
		Folder emailFolder = null;

		try {
//			Session emailSession = Session.getDefaultInstance(properties);
			Session emailSession = Session.getInstance(properties);
			store = emailSession.getStore(protocol);
			store.connect(host2, username2, password2);
			getAllFolders(store);//print all the available folders
//			emailFolder = store.getFolder("INBOX");
			// If want found, try alternative paths		
			
		    emailFolder = store.getFolder("Sent Mail");		    
		    // If still not found, try other common variations
		    if (!emailFolder.exists()) {
		        emailFolder = store.getFolder("[Gmail]/Sent Mail");
		    }
			emailFolder.open(Folder.READ_ONLY);

			Message[] messages = emailFolder.getMessages();
			System.out.println("Total Messages in INBOX: " + messages.length);

			// Fix date format to match email subject format
			SimpleDateFormat sdf = new SimpleDateFormat("d MMM yyyy"); // 'd' avoids leading zero
			String todayDate = sdf.format(new Date());

			// Subject pattern for DSDP data emails from qanawat
			Pattern pattern = Pattern.compile("(?i).*DSDP\\s*data.*");

			
			int maxEmailsToCheck = maxEmailsToChecks;; // Check last 50 emails
			int start = Math.max(0, messages.length - maxEmailsToCheck);

			for (int i = messages.length - 1; i >= start; i--) {
				Message message = messages[i];

				if (!emailFolder.isOpen()) {
					System.out.println("Folder closed unexpectedly.");
					break;
				}

				Date sentDate;
				String subject = message.getSubject();
				System.err.println("Checking Subject: " + subject);

				try {
					sentDate = message.getSentDate();
				} catch (MessagingException e) {
//					System.out.println("Failed to fetch sent date. Skipping message.");
					continue;
				}
				System.out.println("sentDate"+sentDate);
				if (sentDate == null || !isToday(sentDate)) {
					System.out.println("continue sentDate"+sentDate);
					continue; // Skip if not today's mail
				}

				// Match subject with regex
				Matcher matcher = pattern.matcher(subject);
				if (!matcher.find()) {
//					System.out.println("matcher subject"+subject);
					continue; // Skip if subject doesn't match
				}

				// Process multipart email
				if (message.isMimeType("multipart/*")) {
					Multipart multipart = (Multipart) message.getContent();
//					int totalExcelAttachments = multipart.getCount();					
//					int totalExcelAttachments=countValidAttachmentInExcel(multipart,totalParts);					 
//					System.out.println("Total Attachment in mail "+totalExcelAttachments);
					System.out.println("multipart found"+subject);					
					for (int j = 0; j < multipart.getCount(); j++) {
						String generateAlphaCode=HelperExtension.generateShortCode(3);
					    BodyPart bodyPart = multipart.getBodyPart(j);
					    String fileName = j+"_"+generateAlphaCode+"_"+bodyPart.getFileName();
					    System.out.println("filename "+fileName);
					    String disposition = bodyPart.getDisposition();
					    if (!mailProcessed &&
					    (disposition == null || Part.ATTACHMENT.equalsIgnoreCase(disposition)) &&
					        fileName != null &&
					        fileName.toLowerCase().endsWith(".xlsx") &&
					        fileName.toLowerCase().contains(excelSheetName)) {					    	
					        File savedFile = new File(excelPath + fileName);
					        System.out.println("Downloading Attachment: " + fileName);

					        try (InputStream is = bodyPart.getInputStream();
					             OutputStream os = new FileOutputStream(savedFile)) {
					            byte[] buffer = new byte[4096];
					            int bytesRead;
					            while ((bytesRead = is.read(buffer)) != -1) {
					                os.write(buffer, 0, bytesRead);
					            }
					        }

					        System.out.println("Saved File: " + savedFile.getAbsolutePath());

					        try (FileInputStream fis = new FileInputStream(savedFile);
					             Workbook workbook = new XSSFWorkbook(fis)) {
					            saveExcelData(workbook);
					        }

					        System.out.println("Attachment processed successfully.");					        
					        
				            // ✅ Mark mail processed after last Excel attachment only
//				            if (processedCount == totalExcelAttachments) {
//					            mailProcessed = true;
//					            System.out.println("📬 Last attachment processed. Mail marked as processed.");
//						        break; // ✅ Process all the file's in attachement
//					        }

					    }
					}
					 mailProcessed = true;
			        System.out.println("📬 Last attachment processed. Mail marked as processed.");

				}

				if (mailProcessed)
					break; // ✅ Stop checking emails if one is processed
			}

		} catch (FolderClosedException e) {
			System.out.println("Email folder was closed unexpectedly. " + e.getMessage());
		} catch (MessagingException e) {
			System.out.println("Messaging Exception: " + e.getMessage());
		} catch (IOException e) {
			System.out.println("IO Exception: " + e.getMessage());
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (emailFolder != null && emailFolder.isOpen()) {
					emailFolder.close(false);
				}
				if (store != null && store.isConnected()) {
					store.close();
				}
			} catch (MessagingException e) {
				e.printStackTrace();
			}
		}

		if (!mailProcessed) {
			sendNoMailAlert(username2, password2);
			System.err.println("Process finished succesfully----------");
		}
	}
	
	private void getAllFolders(Store store) throws MessagingException {
		Folder defaultFolder = store.getDefaultFolder();
		Folder[] folders = defaultFolder.list("*");

		System.out.println("Available folders:");
		for (Folder folder : folders) {
		    System.out.println("- " + folder.getFullName());
		    
		    // List subfolders
		    if ((folder.getType() & Folder.HOLDS_FOLDERS) != 0) {
		        Folder[] subFolders = folder.list();
		        for (Folder subFolder : subFolders) {
		            System.out.println("  └─ " + subFolder.getFullName());
		        }
		    }
		}
	}
	private int countValidAttachmentInExcel(Multipart multipart,int totalParts) {
		 // Count only valid Excel attachments first
		
		int totalExcelAttachments = 0;
		try {						    
	    for (int j = 0; j < totalParts; j++) {
	        BodyPart part = multipart.getBodyPart(j);
	        String fileName = part.getFileName();
	        if (fileName != null &&
	            Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition()) &&
	            fileName.toLowerCase().endsWith(".xlsx") &&
	            fileName.toLowerCase().contains(excelSheetName)) {
	            totalExcelAttachments++;
	        }
	    }
		}catch(Exception e) {
			e.printStackTrace();			
		}
		
	    return totalExcelAttachments;
	}

	private void sendNoMailAlert(String fromEmail, String password) {
		DayOfWeek today = LocalDate.now().getDayOfWeek();

		if (today == DayOfWeek.SATURDAY || today == DayOfWeek.SUNDAY) {
			System.out.println("No email will be sent today as it is the weekend.");
			return;
		}
		Properties props = new Properties();
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.host", host);
		props.put("mail.smtp.port", port);

		Session session = Session.getInstance(props, new javax.mail.Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(fromEmail, password);
			}
		});

		try {
			Message message = new MimeMessage(session);
			message.setFrom(new InternetAddress(fromEmail));
			message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(email));
			message.setSubject("Alert: DSDP data not received today");
			message.setText("Dear Team,\n\n"
					+ "This is to inform you that the scheduled report email for the client **QANAWAT**, titled **\"DSDP data\"** for today has **not been received**.\n\n"
					+ "Purpose: This report is crucial for monitoring .\n\n"
					+ "Kindly review the status and share the report at the earliest to avoid any operational delays.\n\n"
					+ "Thank you.\n\n" + "Best regards,\n" + "Tech Team");

			Transport.send(message);
			System.out.println("Alert email sent successfully.");

		} catch (MessagingException e) {
			e.printStackTrace();
		}
	}

	private boolean isToday(Date date) {
		if (date == null)
			return false;
		LocalDate mailDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();

		return LocalDate.now().equals(mailDate);
	}
	private void printHeaderFromExcelSheet(Sheet sheet) {
		try {
			Row headerRow = sheet.getRow(0);
			if (headerRow != null) {
			    // Loop through all cells in the header row
			    for (int cellNum = 0; cellNum < headerRow.getLastCellNum(); cellNum++) {
			        String header = getCellValue(headerRow.getCell(cellNum));
			        System.out.print(header + " | ");
			    }
			    System.out.println(); // New line after headers
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
	}

	@Transactional
	public void saveExcelData(Workbook workbook) {
	    try {
	        Sheet sheet = workbook.getSheetAt(0);
	        List<TblDsdpBillingTransaction> records = new ArrayList<>();
	        //print header's 
	        printHeaderFromExcelSheet(sheet);
	        // Skip header row, start from row 1
	        for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
	            Row row = sheet.getRow(rowNum);
	            if (row == null) continue;

	            String timestamp = getCellValue(row.getCell(0));
	            String serviceId = getCellValue(row.getCell(1));
	            String productId = getCellValue(row.getCell(2));
	            String msisdn = getCellValue(row.getCell(3));
	            String feeStr = getCellValue(row.getCell(4));

	            if (timestamp.isEmpty() && serviceId.isEmpty() && msisdn.isEmpty()) continue;

	            double fee = 0.0;
	            try {
	                fee = Double.parseDouble(feeStr);
	            } catch (NumberFormatException ignored) {}

	            TblDsdpBillingTransaction record = TblDsdpBillingTransaction.builder()
	                    .timestamp(timestamp)
	                    .serviceId(serviceId)
	                    .productId(productId)
	                    .msisdn(msisdn)
	                    .fee(fee)
	                    .processedDate(LocalDateTime.now())
	                    .build();

	            records.add(record);
	        }

	        if (!records.isEmpty()) {
	            excelTransactionRepository.saveAll(records);
	            System.out.println("✅ Successfully saved " + records.size() + " records to database.");
	        } else {
	            System.out.println("⚠️ No valid records found in Excel.");
	        }

	    } catch (Exception e) {
	        e.printStackTrace();
	        throw new RuntimeException("Error while saving Excel data", e);
	    }
	}

	private String getCellValue(Cell cell) {
		if (cell == null)
			return "";

		CellType cellType = cell.getCellType();
		if (cellType == CellType.FORMULA) {
			cellType = cell.getCachedFormulaResultType(); // Get actual value type
		}

		switch (cellType) {
		case STRING:
			return cell.getStringCellValue().trim();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				return new SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue());
			}
			return String.valueOf(cell.getNumericCellValue());
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		case BLANK:
		default:
			return "";
		}
	}

	private double parseDouble(String value) {
		if (value == null || value.trim().isEmpty())
			return 0.0;
		try {
			// Remove commas and trim the value
			return Double.parseDouble(value.replace(",", "").trim());
		} catch (NumberFormatException e) {
			e.printStackTrace(); // Optional: log exact cell that failed
			return 0.0;
		}
	}
	
	public String processAllExcelFiles(String folderPath) {
		if(folderPath.isBlank() || folderPath.isEmpty() || folderPath == null) {
			System.out.println("Enter the folder Path");
			return "Enter the folder Path"; 
		}
	    File folder = new File(folderPath);
	    if (!folder.exists() || !folder.isDirectory()) {
	        System.out.println("❌ Invalid folder path: " + folderPath);
	        return "❌ Invalid folder path: " + folderPath;
	    }

	    // Get all .xlsx files
	    File[] excelFiles = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".xlsx"));
	    if (excelFiles == null || excelFiles.length == 0) {
	        System.out.println("⚠️ No Excel files found in folder: " + folderPath);
	        return "⚠️ No Excel files found in folder: " + folderPath;
	    }

	    System.out.println("📂 Found " + excelFiles.length + " Excel files. Starting import...");

	    int processedCount = 0;
	    for (File excelFile : excelFiles) {
	        System.out.println("\n📘 Processing file: " + excelFile.getName());
	        try (FileInputStream fis = new FileInputStream(excelFile);
	             Workbook workbook = new XSSFWorkbook(fis)) {

	            // Your existing DB save logic
	            saveExcelData(workbook);

	            processedCount++;
	            System.out.println("✅ Successfully processed: " + excelFile.getName());

	        } catch (IOException e) {
	            System.err.println("❌ Failed to process file: " + excelFile.getName());
	            e.printStackTrace();
	        }
	    }

	    System.out.println("\n✅ All files processed. Total successful: " + processedCount + "/" + excelFiles.length);
	    return "\n✅ All files processed. Total successful: " + processedCount + "/" + excelFiles.length;
	}
	

}
