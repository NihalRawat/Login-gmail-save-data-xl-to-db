package com.atpl.entity;

import java.time.LocalDateTime;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Table;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Entity
@Table(name = "tbl_dsdp_billing_transaction")
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class TblDsdpBillingTransaction {

	@Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    private String timestamp;
    private String serviceId;
    private String productId;
    private String msisdn;
    private Double fee;
    private LocalDateTime processedDate;
}
