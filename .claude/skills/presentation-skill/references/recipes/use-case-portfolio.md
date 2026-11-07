# Recipe: AI Use-Case Portfolio Slides

## Value vs feasibility (matrix-2x2)

```
## Slide 3: Four of eleven use cases clear the value-feasibility bar
**Layout:** matrix-2x2
- X-axis: Feasibility
- Y-axis: Business value
- Q1: Quick wins
- Q2: Strategic bets
- Q3: Deprioritize
- Q4: Fill-ins
- Item: Name="Claims triage" X="0.8" Y="0.85"
- Item: Name="KYC copilot" X="0.75" Y="0.7"
- Item: Name="Churn models" X="0.6" Y="0.8"
- Item: Name="Contract review" X="0.7" Y="0.65"
- Item: Name="HR chatbot" X="0.85" Y="0.3"
- Item: Name="Full underwriting automation" X="0.2" Y="0.9"
- Notes: Top-right quadrant is the 90-day portfolio.
- Source: Use-case scoring workshop, Aug 2026
```

## 90-day roadmap (gantt)

```
## Slide 9: Two quick wins ship inside 90 days; foundations run in parallel
**Layout:** gantt
**Periods:** Sep, Oct, Nov, Dec
- Bar: Row="Quick wins" Label="Claims triage MVP" Start="1" End="2"
- Bar: Row="Quick wins" Label="KYC copilot pilot" Start="2" End="4"
- Bar: Row="Foundations" Label="Eval harness" Start="1" End="3"
- Bar: Row="Foundations" Label="Governance stand-up" Start="1" End="2"
- Milestone: Row="Quick wins" Label="Steering go/no-go" At="2"
- Notes: Go/no-go gates the Q4 spend.
- Source: Delivery plan v3
```

`**Periods:**` is a deck-level bold field (not a `- Periods:` bullet); `Bar:`
rows use `Row=`/`Label=`/`Start=`/`End=`, and `Milestone:` rows use
`Row=`/`Label=`/`At=` — mirror `references/generation-guide.md`'s gantt
section exactly before adapting.
