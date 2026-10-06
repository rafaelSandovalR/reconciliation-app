package com.apex.reconciliation_app.repository;

import com.apex.reconciliation_app.model.shopsimon.ShopSimonRawTransaction;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface ShopSimonRawTransactionRepository extends JpaRepository<ShopSimonRawTransaction, Long> {
    boolean existsByCompositeTransactionId(String compositeTransactionId);
}
