package com.apex.reconciliation_app.repository;

import com.apex.reconciliation_app.model.TemuRawTransaction;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface TemuRawTransactionRepository extends JpaRepository<TemuRawTransaction, Long> {
    boolean existsByCompositeTransactionId(String compositeTransactionId);
}
