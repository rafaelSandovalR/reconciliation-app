package com.apex.reconciliation_app.dto;

import java.util.List;

public record MarketplaceParseResult<S, A>(
        List<S> errorSuspense,
        List<S> actionableSuspense,
        List<A> auditTrail,
        List<String> consoleLogs
) {}
