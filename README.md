# Third-Party Risk Assessment

## Project Description

This project assesses eight fictional vendors the way a third-party risk management (TPRM) team would. Each vendor is tiered by inherent risk, scored on a 12-question weighted due diligence questionnaire, checked against its SOC 2 report where one is required, and given a residual risk rating, a next reassessment date and an onboarding or renewal decision. The scoring is formula-driven, so changing an answer, a weight or a threshold updates every rating. It also puts the finding from my risk register (critical vendors not reassessed in over 18 months) into practice, and it supports the supply chain category in my [`nist-csf-gap-assessment`](https://github.com/mmanzanares7/nist-csf-gap-assessment). This is the kind of work GRC teams do to manage vendor risk before and after a contract is signed.

## Files

| File | Purpose |
|---|---|
| [`vendor_risk_assessment.xlsx`](vendor_risk_assessment.xlsx) | Workbook: Scope, Tiering, Questionnaire, SOC2 Review, Decisions and Scale sheets (201 formulas, no errors) |
| [`vendor_decisions_export.csv`](vendor_decisions_export.csv) | The decisions table, readable on GitHub |
| [`scripts/build_vendor_assessment.py`](scripts/build_vendor_assessment.py) | Python (openpyxl) script that generates the workbook |

## How it works

1. **Tiering:** each vendor is scored 1 to 3 on data sensitivity (35%), system access (25%), business criticality (25%) and replaceability (15%). The weighted score sets Tier 1 (2.4 or more), Tier 2 (1.5 or more) or Tier 3, and the tier sets the reassessment cycle (12, 24 or 36 months) and the diligence required.
2. **Questionnaire:** 12 weighted questions cover governance, assurance, access, data protection, operations, change, incident response, continuity, subcontractors and training. Yes counts fully, Partial counts half, and N/A questions are excluded. The control score is the weighted share earned, rated Strong (85% or more), Adequate (70% or more) or Weak.
3. **SOC 2 review:** for each report the workbook calculates the report's age, flags it as stale after 12 months, counts user entity controls not yet mapped to our own controls, and writes a conclusion: acceptable, request a bridge letter, follow up on exceptions, map controls, or escalate.
4. **Decisions:** residual risk comes from a tier by control strength matrix (for example, Tier 1 with Weak controls is Critical). The workbook then sets the next due date with `EDATE`, flags overdue reassessments and writes the decision.

## Results

| Vendor | Tier | Control score | Residual risk | SOC 2 | Reassessment | Decision |
|---|---|---|---|---|---|---|
| Summit IT Services | 1 | 53% (Weak) | Critical | Required, not provided | Overdue | Do not onboard or renew until critical gaps are fixed |
| Brightpath Software | 1 | 72% (Adequate) | High | 14 months old, 4 user controls unmapped | Overdue | Approve with conditions |
| Pinecrest Staffing | 2 | 43% (Weak) | High | Not required | Overdue | Approve with conditions |
| Nimbus Cloud Hosting | 1 | 91% (Strong) | Moderate | Acceptable | Current | Approve and monitor |
| Keystone Legal LLP | 2 | 80% (Adequate) | Moderate | Not required | Current | Approve and monitor |
| Oakmont Audit Partners | 2 | 94% (Strong) | Low | Not required | Current | Approve |
| Lighthouse Training Co | 3 | 83% (Adequate) | Low | Not required | Current | Approve |
| Harvest Food Service | 3 | Not assessed | Not assessed | Not required | Current | Approve (attestation only) |

- **Summit IT Services** is the one decision that needs escalation. It has privileged access to servers but no independent security audit, no multi-factor authentication and partial offboarding controls, and its last reassessment was 21 months ago.
- **Brightpath Software** scores Adequate on controls, but its SOC 2 report ends 14 months before the as-of date, so the assessment asks for a bridge letter covering the gap, and 4 of 10 user entity controls are not yet mapped to our own controls. Those are the controls the vendor expects us to perform.
- **Three reassessments are overdue**, two of them Tier 1. These are the vendors the risk register describes as critical and not reassessed in over 18 months.
- **Tier 3 vendors** get a lighter review on purpose. Harvest Food Service has no questionnaire at all, because requiring one would spend effort where no company data is involved.

## Connections to other projects

- [`risk-register-excel`](https://github.com/mmanzanares7/risk-register-excel) lists the vendor reassessment gap as R-15.
- [`nist-csf-gap-assessment`](https://github.com/mmanzanares7/nist-csf-gap-assessment) rates GV.SC (supply chain risk management) as an immediate priority and recommends this process.

## Limitations

All vendors, answers and dates are fictional, and the factor scores and thresholds are my judgments. Questionnaire answers are self-reported here. A real program verifies high-risk answers with evidence such as policies, screenshots and penetration test summaries, and reads the full SOC 2 report, including the auditor's description of tests and any exceptions, not only the opinion.

## Summary

I built a repeatable vendor risk process that moves from tiering to diligence to a decision, with every score traceable to an input and every overdue item flagged automatically. It shows how I would size diligence to risk, read a SOC 2 report for what the user has to do, and turn findings into decisions a procurement or compliance team can act on.
