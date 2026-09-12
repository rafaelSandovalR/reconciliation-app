package com.apex.reconciliation_app.repository;

import com.apex.reconciliation_app.model.EbayRawTransaction;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface EbayRawTransactionRepository extends JpaRepository<EbayRawTransaction, Long> {
    boolean existsByCompositeTransactionId(String compositeTransactionId);
}
