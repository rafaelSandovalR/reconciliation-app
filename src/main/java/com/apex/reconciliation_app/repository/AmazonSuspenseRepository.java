package com.apex.reconciliation_app.repository;

import com.apex.reconciliation_app.model.AmazonSuspense;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface AmazonSuspenseRepository extends JpaRepository<AmazonSuspense, Long> {
}
