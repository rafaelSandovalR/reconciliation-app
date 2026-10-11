package com.apex.reconciliation_app.repository;

import com.apex.reconciliation_app.model.tiktok.TikTokSuspense;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface TikTokSuspenseRepository extends JpaRepository<TikTokSuspense, Long> {
}
