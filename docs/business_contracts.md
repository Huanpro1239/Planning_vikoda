# Planning Business Contracts

This document defines the production-planning invariants that refactors must
preserve. The executable source of truth is
`tests/contracts/test_planning_business_invariants.py`.

| Contract | Invariant | Observable contract |
|---|---|---|
| CONTRACT-01 | Service First | Sales/mandatory service is scheduled before safety-stock buffer. Service completion may still be publish-ready when only buffer is short. |
| CONTRACT-02 | Debt priority | Debt is mandatory service. With FC=0 and debt>0, production starts on day 1. A debt SKU wins equal-time priority on the shared machine. |
| CONTRACT-02B | Debt formula mode | `SUBTRACT_BOOK_ON_DEBT` deducts book stock; `IGNORE_BOOK_ON_DEBT` does not. Debt branch does not add target ending stock. |
| CONTRACT-03 | Safety Stock | Safety stock uses residual capacity only and can never displace required service. Workbook P/Q/S+ reflect physically schedulable quantity, not unreachable buffer. |
| CONTRACT-04 | KHS/PET9000 shared machine | KHS and PET 9000 consume one serialized capacity pool and cannot exceed shared shifts/day. |
| CONTRACT-04B | Changeover | Mold change consumes setup capacity and reduces productive capacity in the affected day. |
| CONTRACT-05 | Publish boundary | Ready plans auto-authorize; stockout-risk plans require exact proposal approval with reason; stale approvals are rejected. |
| CONTRACT-05B | Non-waivable validation | Resource validation failure cannot be overridden by review approval. |

## Test boundary

The suite deliberately tests public behavior across several layers:

```text
weekly_engine
    ↓
weekly_model
    ↓
workbook + report verification
    ↓
publish proposal identity
    ↓
publish policy
```

The tests do not assert private implementation structure. Refactors may move
functions or change internal algorithms as long as the observable business
contracts remain identical.

## Change policy

A contract test must not be weakened merely to make a refactor pass. If a
business rule is intentionally changed, the change should be reviewed as a
business-policy change and the contract/documentation updated in the same PR
with an explicit explanation of the new rule.
