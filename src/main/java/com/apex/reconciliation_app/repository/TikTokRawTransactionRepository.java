package com.apex.reconciliation_app.repository;

import com.apex.reconciliation_app.model.tiktok.TikTokRawTransaction;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface TikTokRawTransactionRepository extends JpaRepository<TikTokRawTransaction, Long> {
    boolean existsByCompositeTransactionId(String compositeTransactionId);
}
