# Auto-Create Task for High-Value Opportunities

## Business Problem
Sales reps may forget to create follow-up tasks for high-value Opportunities, leading to inconsistent follow-up and missed executive visibility.

## Solution
Enhanced a Salesforce Record-Triggered Flow to automatically create different Tasks based on Opportunity Amount.

- $10,000–$49,999 creates a Standard High-Value Review Task.
- $50,000+ creates an Executive Review Required Task.

## Business Value
- Reduces manual task creation
- Improves sales process consistency
- Provides executive visibility into major deals
- Prevents unnecessary duplicate flows
- Centralizes related automation in one maintainable Flow

## Key Salesforce Features
- Record-Triggered Flow
- After-Save Flow
- Decision Element
- Create Records
- Task Automation
- Dynamic field assignment

## What I Learned
- How to use Decision elements for branching logic
- How to avoid overlapping automations
- Difference between Assigned To ID, Related To ID, WhatId, and WhoId
- Why Flow paths do not always need to reconnect
- Why updated-record entry settings matter

## Testing Notes
- $9,500 to $15,000 created the standard high-value task.
- $15,000 to $60,000 did not create the executive task under the original entry setting because the record already met the $10,000 entry condition.
- Next improvement: update trigger behavior and add duplicate-prevention logic.