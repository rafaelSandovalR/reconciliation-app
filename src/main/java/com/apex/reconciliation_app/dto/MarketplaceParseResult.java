package com.apex.reconciliation_app.dto;

import java.util.List;

public record MarketplaceParseResult<S, A>(
        List<S> suspenseQueue,
        List<A> auditTrail
) {}
