package com.apex.reconciliation_app.repository;

import com.apex.reconciliation_app.model.ReconciliationRecord;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface ReconciliationRepository extends JpaRepository<ReconciliationRecord, String> {

}
