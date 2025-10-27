package com.atpl.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.atpl.entity.TblExcelTransaction;

@Repository
public interface ExcelTransactionRepository extends JpaRepository<TblExcelTransaction, Long>{

}
