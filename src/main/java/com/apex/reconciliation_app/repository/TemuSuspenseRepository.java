package com.apex.reconciliation_app.repository;

import com.apex.reconciliation_app.model.TemuSuspense;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface TemuSuspenseRepository extends JpaRepository<TemuSuspense, Long> {
}
