package com.apex.reconciliation_app.repository;

import com.apex.reconciliation_app.model.shopsimon.ShopSimonSuspense;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface ShopSimonSuspenseRepository extends JpaRepository<ShopSimonSuspense, Long> {
}
