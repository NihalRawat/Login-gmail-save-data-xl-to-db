package com.atpl.utility;

import java.security.SecureRandom;

import org.springframework.stereotype.Component;

@Component
public class HelperExtension {
	public static String generateShortCode(int length) {
	    String chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
	    SecureRandom random = new SecureRandom();
	    StringBuilder sb = new StringBuilder(length);
	    for (int i = 0; i < length; i++) {
	        sb.append(chars.charAt(random.nextInt(chars.length())));
	    }
	    return sb.toString();
	}
	
}
