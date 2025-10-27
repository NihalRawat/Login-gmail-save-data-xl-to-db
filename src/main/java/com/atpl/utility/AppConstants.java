package com.atpl.utility;

public class AppConstants {

	public static final Integer SuccessCode = 200;
	public static final String SuccessMessage = "Success";
	public static final Integer BadReqCode = 400;
	public static final String BadReqMessage = "Bad Request: ";
	public static final String InvalidCredentialsMessage = "Invalid credentials";
	public static final Integer NotFoundCode = 404;
	public static final String NotFoundMessage = "Not Found: ";
	public static final Integer DuplicateCode = 409;
	public static final Integer PASSWORD_POLICY_VIOLATION = 422;
	public static final String PASSWORD_POLICY_MESSAGE = "Password does not meet security requirements.";
	public static final Integer DuplicatePasswordCode = 424;
	public static final String DuplicatePasswordMessage = "New password must be different from the old password.";
	public static final String DuplicateUserMessage = "User already exists: ";
	public static final Integer InternalServerErrorCode = 500;
	public static final String InternalServerErrorMessage = "Something went wrong.";
	public static final String UserNotFound = "User not found with username: ";
	public static final String AccountLockedMessage = "Account locked. please contact admin.";
	public static final String UploadCodSlipMessage = "COD slip uploaded successfully";
	public static final String InvalidFileFormatCode="Invalid file format. Only JPG, JPEG, PNG, BMP files are allowed";
	public static final String DuplicatePaymentSlip = "A payment slip has already been uploaded for this transaction with the same details.";
	public static final Integer Password_Reset_Link_Sent_To_Your_Mail = 567;
	public static final String Password_Reset_Link = "Password Reset Link is sent on Mail";
	public static final Integer Email_Id_Not_Correct_Code = 568;
	public static final Integer Email_Id_Not_Exist_Code = 569;
	public static final String Email_Id_Not_Correct_Message ="Email Id Is Not Correct";
	public static final String Log_Out_Message="You have successfully logged out";
}
