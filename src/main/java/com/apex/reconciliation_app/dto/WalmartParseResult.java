package com.apex.reconciliation_app.dto;

import com.apex.reconciliation_app.model.WalmartRawTransaction;
import com.apex.reconciliation_app.model.WalmartSuspense;

import java.util.List;

public record WalmartParseResult(
        List<WalmartSuspense> suspenseQueue,
        List<WalmartRawTransaction> auditTrail
) {}
