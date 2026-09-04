package com.apex.reconciliation_app.repository;

import com.apex.reconciliation_app.model.ReconciliationRecord;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface ReconciliationRepository extends JpaRepository<ReconciliationRecord, String> {
    List<ReconciliationRecord> findBySiteOrderId(String siteOrderId);
}
