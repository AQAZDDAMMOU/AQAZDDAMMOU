# Recipe: AI Governance Deck Slides

Copy-adapt these validated slide blocks. Pair with `references/storyline.md`
(SCQA the exec summary first) and action titles throughout.

## Governance model (org-chart)

```
## Slide 4: Governance concentrates decisions in one accountable forum
**Layout:** org-chart
- Node: Id="board" Label="Board Risk Committee" Role="Oversight"
- Node: Id="steer" Label="AI Steering Committee" Parent="board" Role="Direction & funding"
- Node: Id="coe" Label="AI Center of Excellence" Parent="steer" Role="Delivery & standards"
- Node: Id="mrisk" Label="Model Risk Mgmt" Parent="steer" Role="2nd line challenge"
- Node: Id="owners" Label="Business Model Owners" Parent="steer" Role="1st line accountability"
- Node: Id="audit" Label="Internal Audit" Parent="board" Role="3rd line assurance"
- Notes: Three lines of defense; steering committee meets monthly.
- Source: Operating model design workshop, Aug 2026
```

## Decision rights (raci)

```
## Slide 5: Decision rights are unambiguous across the model lifecycle
**Layout:** raci

| Activity | Business | AI CoE | Model Risk | Audit |
|---|---|---|---|---|
| Use-case approval | A | R | C | I |
| Model deployment | C | R | A | I |
| Incident response | R | R | A | I |

- Notes: A single Accountable per row — argue about the A's, not the C's.
- Source: RACI workshop output, Aug 2026
```

## Escalation flow (swimlane-process)

```
## Slide 6: Incidents escalate through two gates in under 48 hours
**Layout:** swimlane-process
- Lane: Name="Model Owner" Steps="Detect drift > Triage severity"
- Lane: Name="AI CoE" Steps="Contain > Root-cause"
- Lane: Name="Model Risk" Steps="Assess materiality > Approve remediation"
- Notes: Severity-1 skips triage and goes straight to Model Risk.
- Source: Incident runbook v2
```
