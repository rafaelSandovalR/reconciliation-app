package com.apex.reconciliation_app.repository;

import com.apex.reconciliation_app.model.WalmartRawTransaction;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface WalmartRawTransactionRepository extends JpaRepository<WalmartRawTransaction, Long> {
    boolean existsByCompositeTransactionId(String compositeTransactionId);
}
