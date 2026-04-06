# Test Case Generation Guide

This project is the **CDS Front Collection** system for an insurance company. Test cases are written in structured markdown files and converted to Excel.

## Project Structure

```
test_cases_by_module/
  <Module Folder>/
    <module_name>.md
convert_to_excel.py       ← regenerates test-cases-all-modules.xlsx
test-cases-all-modules.xlsx
```

Current modules:
- Agent Transactions
- Collection Reversal
- Group Policy
- Others
- Policy Direct Marketing
- Policy Regular (BASE)
- Policy VUL
- Regular (Balance Payment)
- SSI

## Regenerating the Excel

After creating or editing any `.md` file, run:
```
python3 convert_to_excel.py
```

## Markdown File Format

Every test case `.md` file must follow this exact structure:

```markdown
# Test Cases - [Module Name]

## Module Information
- **Module Name:** [Module Name]
- **Total Test Cases:** [count]

# TCD_IL_Sprint5-CDS Front Collection

## Test Case Template Columns
- Test Case #
- User Story
- Scenario
- Preconditions
- Expected Result
- Remarks/Documentation
- Test Data

---

## TC-[number]: [Scenario Title]

| Field | Details |
|-------|---------|
| Test Case # | TC-[number] |
| User Story | As [role], I want to [action] |
| Scenario | [Clear description of what is being tested] |
| Preconditions | <br>- [condition 1]<br>- [condition 2] |
| Expected Result | <br>- [result 1]<br>- [result 2] |
| Remarks/Documentation | See acceptance criteria |
| Test Data | [field]: [value] | [field]: [value] |

---

## Module Summary

Total Test Cases in this Module: [count]

---

**End of [Module Name] Test Cases**
```

## Test Case Writing Rules

- TC IDs are sequential per module (TC-001, TC-002, ...) or use a module prefix (TC-CR-001 for Collection Reversal, etc.)
- Always cover: Valid Input, Invalid Input, Edge Cases (zero amount, boundary values), status-based scenarios
- Use realistic test data: Policy No: A12345678, Amount: ₱5,000.00, Teller ID: TLR001, Supervisor ID: SUP001
- Preconditions and Expected Results use `<br>-` prefix for each bullet point
- Group test cases logically: valid → invalid → edge cases → status variations
- Scenario titles should start with "Verify..."
- Each acceptance criteria item should have at least one test case

## Excel Output Columns (per sheet/module)

| Column | Description |
|--------|-------------|
| Module | Module name |
| Test Case ID | TC-XXX |
| User Story | Full user story text |
| Test Case Title | Normalized scenario title |
| Preconditions | Bullet list |
| Steps | Auto-generated numbered steps with per-step expected result in Expected Results column |
| Expected Results | One row per step result (each in its own cell) |
| Test Data | Raw test data string |
| Status | Not Started / Pass / Fail / In Progress |

## Key Business Context

- System: CDS Front Collection
- Users: Teller, Supervisor
- Policy sources: Ingenium, SUKI, AUS (searched in sequence)
- TranCodes identify transaction types (e.g., 1004 = renewal premium)
- LOB (Line of Business) is determined by policy number prefix
- OR = Official Receipt, generated on every successful transaction
- Supervisor override is required for sensitive operations (e.g., reversals)
- GL accounts are reversed on collection reversal
- Same-day rule applies for reversals
