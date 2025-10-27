//package com.atpl.service;
//
//import io.jsonwebtoken.Claims;
//import io.jsonwebtoken.Jwts;
//import io.jsonwebtoken.SignatureAlgorithm;
//import org.springframework.beans.factory.annotation.Value;
//import org.springframework.stereotype.Service;
//
//import java.util.Date;
//import java.util.HashMap;
//import java.util.Map;
//import java.util.function.Function;
//
//@Service
//public class JwtService {
//
//	@Value("${jwt.expire-time}")
//	private long expireTime;
//	@Value("${jwt.secret}")
//	private String jwtSecret;
//
//	public String generateToken(String userName) {
//		Map<String, Object> claims = new HashMap<>();
//		return createToken(claims, userName);
//	}
//
//	private String createToken(Map<String, Object> claims, String subject) {
//		return Jwts.builder().setClaims(claims).setSubject(subject).setIssuedAt(new Date(System.currentTimeMillis()))
//				.setExpiration(new Date(System.currentTimeMillis() + expireTime))
//				.signWith(SignatureAlgorithm.HS256, jwtSecret).compact();
//	}
//
//	public Boolean validateToken(String token, String userName) {
//		try {
//			String username = extractUsername(token);
//			return (username.equals(userName) && !isTokenExpired(token));
//		} catch (Exception e) {
//			return false;
//		}
//	}
//
//	private Boolean isTokenExpired(String token) {
//		return extractExpiration(token).before(new Date());
//	}
//
//	private Claims extractAllClaims(String token) {
//		return Jwts.parser().setSigningKey(jwtSecret).parseClaimsJws(token).getBody();
////		try {
////			return Jwts.parser().setSigningKey(jwtSecret).parseClaimsJws(token).getBody();
////		} catch (Exception e) {
////			throw new RuntimeException("Invalid JWT Token", e);
////		}
//	}
//
//	public String extractUsername(String token) {
//		return extractClaim(token, Claims::getSubject);
//	}
//
//	public Date extractExpiration(String token) {
//		return extractClaim(token, Claims::getExpiration);
//	}
//
//	public <T> T extractClaim(String token, Function<Claims, T> claimsResolver) {
//		final Claims claims = extractAllClaims(token);
//		return claimsResolver.apply(claims);
//	}
//
//	public String generateToken(String userName, Map<String, Object> additionalClaims) {
//		Map<String, Object> claims = new HashMap<>(additionalClaims);
//		return createToken(claims, userName);
//	}
//}
