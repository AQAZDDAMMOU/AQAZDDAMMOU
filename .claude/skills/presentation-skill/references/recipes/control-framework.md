# Recipe: AI Control Framework Slides

## Lifecycle x control-type matrix (heatmap-table)

Lifecycle stages across the top; control families down the side; cell value
= coverage maturity (renders as a colored heatmap).

```
## Slide 7: Controls thin out exactly where model risk peaks — in production
**Layout:** heatmap-table

| Control family | Design | Build | Validate | Deploy | Monitor |
|---|---|---|---|---|---|
| Policy & standards | 4 | 4 | 3 | 2 | 1 |
| Evaluation & testing | 3 | 4 | 4 | 2 | 1 |
| Human oversight | 4 | 3 | 3 | 2 | 2 |
| Documentation | 4 | 3 | 3 | 3 | 2 |
| Incident response | 2 | 2 | 2 | 1 | 1 |

- Notes: Read the right-hand columns — monitoring is the systemic gap.
- Source: Control self-assessment, Q3 2026
```

## Control operating layers (architecture-layers)

```
## Slide 8: Five control layers wrap every production model
**Layout:** architecture-layers
- Layer: Name="Oversight" Items="Steering committee, Board reporting"
- Layer: Name="Policy" Items="AI policy, Use standards, Risk tiers"
- Layer: Name="Evaluation" Items="Pre-deploy evals, Red-teaming" Accent="2"
- Layer: Name="Monitoring" Items="Drift, Incidents, Usage"
- Layer: Name="Foundation" Items="Access control, Lineage, Audit trail"
- Notes: Accent layer = this quarter's investment focus.
- Source: Target control architecture, v1
```
