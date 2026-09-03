package com.apex.reconciliation_app.repository;

import com.apex.reconciliation_app.model.AmazonRawTransaction;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface AmazonRawTransactionRepository extends JpaRepository<AmazonRawTransaction, Long> {
    boolean existsByCompositeTransactionId(String compositeTransactionId);
}
