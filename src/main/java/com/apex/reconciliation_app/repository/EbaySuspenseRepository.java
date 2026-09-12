package com.apex.reconciliation_app.repository;

import com.apex.reconciliation_app.model.EbaySuspense;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface EbaySuspenseRepository extends JpaRepository<EbaySuspense, Long> {
}
