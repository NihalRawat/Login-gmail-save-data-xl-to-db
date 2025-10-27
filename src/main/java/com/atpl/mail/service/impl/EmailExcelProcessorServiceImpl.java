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

import com.atpl.entity.TblExcelTransaction;
import com.atpl.mail.service.EmailExcelProcessorService;
import com.atpl.repository.ExcelTransactionRepository;

@Service
public class EmailExcelProcessorServiceImpl implements EmailExcelProcessorService {

	@Value("${EXCEL_STORAGE_PATH}")
	private String excelPath;

	@Autowired
	private ExcelTransactionRepository excelTransactionRepository;

	@Value("${spring.username}")
	private String username2;

	@Value("${spring.password}")
	private String password2;
	@Value("${POP_HOST_VALUE}")
	private String host2;

	@Value("${POP_PORT_VALUE}")
	private String port2;

	@Value("${POP_PROTOCOl}")
	private String protocol;

	@Value("${spring.username}")
	private String username;

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

	@Override
	public void downloadAndProcessAttachment() {
		Properties properties = new Properties();
		properties.put("mail.store.protocol", protocol);
		properties.put("mail.pop3s.host", host2);
		properties.put("mail.pop3s.port", port2);
		properties.put("mail.pop3s.ssl.enable", "true");

		boolean mailProcessed = false;
		Store store = null;
		Folder emailFolder = null;

		try {
			Session emailSession = Session.getDefaultInstance(properties);
			store = emailSession.getStore(protocol);
			store.connect(host2, username2, password2);

			emailFolder = store.getFolder("INBOX");
//			Folder teamMemberFolder = emailFolder.getFolder("Team-Members");
//			Folder gitishFolder = teamMemberFolder.getFolder("Gitish");
			emailFolder.open(Folder.READ_ONLY);

			Message[] messages = emailFolder.getMessages();
			System.out.println("Total Messages in INBOX: " + messages.length);

			// Fix date format to match email subject format
			SimpleDateFormat sdf = new SimpleDateFormat("d MMM yyyy"); // 'd' avoids leading zero
			String todayDate = sdf.format(new Date());
//			String expectedSubject = "ODH COD Pendency " + todayDate + " - DEPENDO";

//Pattern pattern = Pattern.compile(
//					"(?i)\\s*(FW:|Fwd:|RE:)?\\s*ODH COD Pendency \\d{1,2}(st|nd|rd|th)? \\w+ \\d{4} - DEPENDO");
			
			// Subject pattern for DSDP data emails from qanawat
//			Pattern pattern = Pattern.compile("(?i)(FW:|Fwd:)?\\s*\\[dcb\\.support@qanawat-me\\.com\\]\\s*(FW:|Fwd:)?\\s*DSDP data");
			Pattern pattern = Pattern.compile("(?i).*DSDP\\s*data.*");

			
			int maxEmailsToCheck = 50; // Check last 20 emails
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
//					System.out.println("continue sentDate"+sentDate);
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

					System.out.println("multipart found"+subject);
					for (int j = 0; j < multipart.getCount(); j++) {
					    BodyPart bodyPart = multipart.getBodyPart(j);
					    String fileName = bodyPart.getFileName();
					    System.out.println("filename"+fileName);
					    if (!mailProcessed &&
					        Part.ATTACHMENT.equalsIgnoreCase(bodyPart.getDisposition()) &&
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
					        mailProcessed = true;
					        break; // ✅ Exit attachment loop
					    }
					}

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
		}
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


	@Transactional
	public void saveExcelData(Workbook workbook) {
	    try {
	        Sheet sheet = workbook.getSheetAt(0);
	        List<TblExcelTransaction> records = new ArrayList<>();

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

	            TblExcelTransaction record = TblExcelTransaction.builder()
	                    .timestamp(timestamp)
	                    .serviceId(serviceId)
	                    .productId(productId)
	                    .msisdn(msisdn)
	                    .fee(fee)
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

//	@Transactional
//	public void saveExcelDataBackUP(Workbook workbook) {
//	    try {
//	        Sheet sheet = workbook.getSheetAt(0);
//	        Map<String, TblCodDependoReport> excelDataMap = new HashMap<>();
//	        Set<String> hubNamesInSheet = new HashSet<>();
//
//	        for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
//	            Row row = sheet.getRow(rowNum);
//	            if (row == null) continue;
//
//	            String hubName = getCellValue(row.getCell(2));
//	            String collectionDate = getCellValue(row.getCell(10));
//	            String status = getCellValue(row.getCell(27));
//	            if (hubName.isEmpty() || collectionDate.isEmpty()) continue;
//
//	            hubNamesInSheet.add(hubName);
//	            String key = hubName + "|" + collectionDate + "|" + status;
//
//	            double amountToBeCollected = parseDouble(getCellValue(row.getCell(12)));
//	            double collected = parseDouble(getCellValue(row.getCell(13)));
//	            double deposited = parseDouble(getCellValue(row.getCell(14)));
//	            double diff = parseDouble(getCellValue(row.getCell(26)));
//
//	            if (excelDataMap.containsKey(key)) {
//	                TblCodDependoReport existing = excelDataMap.get(key);
//	                existing.setAmountToBeCollected(String.valueOf(parseDouble(existing.getAmountToBeCollected()) + amountToBeCollected));
//	                existing.setCollected(String.valueOf(parseDouble(existing.getCollected()) + collected));
//	                existing.setDeposited(String.valueOf(parseDouble(existing.getDeposited()) + deposited));
//	                existing.setDiff(String.valueOf(parseDouble(existing.getDiff()) + diff));
//	            } else {
//	                TblCodDependoReport report = TblCodDependoReport.builder()
//	                        .srNo(getCellValue(row.getCell(0)))
//	                        .month(getCellValue(row.getCell(1)))
//	                        .hubName(hubName)
//	                        .finalHubName(getCellValue(row.getCell(3)))
//	                        .pickupCode(getCellValue(row.getCell(4)))
//	                        .partnerName(getCellValue(row.getCell(5)))
//	                        .zone(getCellValue(row.getCell(6)))
//	                        .fy(getCellValue(row.getCell(7)))
//	                        .hubType(getCellValue(row.getCell(8)))
//	                        .transactionId(getCellValue(row.getCell(9)))
//	                        .collectionDate(collectionDate)
//	                        .monthAgeing(getCellValue(row.getCell(11)))
//	                        .amountToBeCollected(String.valueOf(amountToBeCollected))
//	                        .collected(String.valueOf(collected))
//	                        .deposited(String.valueOf(deposited))
//	                        .diff(String.valueOf(diff))
//	                        .status(status)
//	                        .lineStatus(getCellValue(row.getCell(28)))
//	                        .bankDate1(getCellValue(row.getCell(29)))
//	                        .ref1(getCellValue(row.getCell(30)))
//	                        .bankDate2(getCellValue(row.getCell(31)))
//	                        .ref2(getCellValue(row.getCell(32)))
//	                        .pmBankDate3(getCellValue(row.getCell(33)))
//	                        .pmRef3(getCellValue(row.getCell(34)))
//	                        .myntraSynergyBankDate1(getCellValue(row.getCell(35)))
//	                        .myntraSynergyRef1(getCellValue(row.getCell(36)))
//	                        .dnDate(getCellValue(row.getCell(37)))
//	                        .dnRef(getCellValue(row.getCell(38)))
//	                        .codToUnidentifiedContraPm(getCellValue(row.getCell(39)))
//	                        .bankDate3(getCellValue(row.getCell(40)))
//	                        .ref3(getCellValue(row.getCell(41)))
//	                        .lostShipmentTrackingId(getCellValue(row.getCell(42)))
//	                        .lostShipmentRecoveryAccountingStatus(getCellValue(row.getCell(43)))
//	                        .interestOnDnCnAccountingStatus(getCellValue(row.getCell(44)))
//	                        .pmCreditAccountingStatus(getCellValue(row.getCell(45)))
//	                        .kiranaToCodAccountingStatus(getCellValue(row.getCell(46)))
//	                        .codToUnidentifiedContraPmAccountingStatus(getCellValue(row.getCell(47)))
//	                        .dnAppliedAccountingStatus(getCellValue(row.getCell(48)))
//	                        .dnReversalAccountingStatus(getCellValue(row.getCell(49)))
//	                        .hubRemarks(getCellValue(row.getCell(50)))
//	                        .hubCategory(getCellValue(row.getCell(51)))
//	                        .depositSlipNumber(getCellValue(row.getCell(52)))
//	                        .depositDate(getCellValue(row.getCell(53)))
//	                        .depositAmount(getCellValue(row.getCell(54)))
//	                        .eklFinRemarks(getCellValue(row.getCell(55)))
//	                        .bankCreditAmountMerged(getCellValue(row.getCell(56)))
//	                        .kiranaAmount(getCellValue(row.getCell(57)))
//	                        .posAmt(getCellValue(row.getCell(58)))
//	                        .kiranaAccountingStatus(getCellValue(row.getCell(59)))
//	                        .accountedTillJul23(getCellValue(row.getCell(60)))
//	                        .myntraSynergyAccountingStatus(getCellValue(row.getCell(61)))
//	                        .finalGrouping(getCellValue(row.getCell(62)))
//	                        .responsibility(getCellValue(row.getCell(63)))
//	                        .spoc(getCellValue(row.getCell(64)))
//	                        .recoverableNonRecoverable(getCellValue(row.getCell(65)))
//	                        .lineOfBusiness(getCellValue(row.getCell(66)))
//	                        .bank(getCellValue(row.getCell(67)))
//	                        .partnerName1(getCellValue(row.getCell(68)))
//	                        .aa(getCellValue(row.getCell(69)))
//	                        .clientName("flipkart")
//	                        .build();
//
//	                excelDataMap.put(key, report);
//	            }
//	        }
//
//	        // Step: Delete all records for hubNames present in Excel sheet
//	        if (!hubNamesInSheet.isEmpty()) {
//	            dependoReportRepository.deleteAllByHubNameIn(new ArrayList<>(hubNamesInSheet));
//	        }
//
//	        // Save new data
//	        dependoReportRepository.saveAll(excelDataMap.values());
//
//	        System.out.println("Excel data import completed successfully.");
//	    } catch (Exception e) {
//	        e.printStackTrace();
//	        throw new RuntimeException("Error while saving Excel data", e);
//	    }
//	}


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

}
