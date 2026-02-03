# Quick Reference Guide - Attendance Justifications (v2.3)

## Dropdown Options & Deductions

### No Deduction (Green)
| Status | Description |
|--------|-------------|
| Normal | Regular attendance |
| Late (Approved) | Approved late arrival |
| Early Departure (Approved) | Approved early leave |
| Annual Leave | Paid annual leave |
| Casual Leave | Casual/personal leave |
| Marriage Leave | Marriage leave |
| Paternity Leave | Paternity leave |
| Maternity Leave | Maternity leave |
| Bereavement Leave | Bereavement leave |
| Military Call Leave | Military service leave |

### With Deduction

| Status | Deduction | Color |
|--------|-----------|-------|
| **Late** | 100/200/500 EGP | Yellow |
| **Absent** | 2 days | Red |
| **Missing Punch In** | 0.5 day (after 3) | Pink |
| **Missing Punch In (Justified)** | 0.5 day (after 3) | Pink |
| **Missing Punch Out** | 0.5 day (after 3) | Pink |
| **Missing Punch Out (Justified)** | 0.5 day (after 3) | Pink |
| **Early Departure** | 0.5 day | Orange |
| **Half Day** | 0.5 day | Orange |
| **Early Leave (HD)** | 0.5 day | Orange |
| **Sick Leave** | 0.25 day | Light Blue |
| **Unpaid Leave** | 1 day | Light Blue |

### Backdated Leaves (BD) - Purple
For leaves transferred from previous months:

| Status | Deduction |
|--------|-----------|
| **Annual Leave (BD)** | No deduction |
| **Casual Leave (BD)** | No deduction |
| **Sick Leave (BD)** | 0.25 day |
| **Unpaid Leave (BD)** | 1 day |
| **Half Day (BD)** | 0.5 day |
| **Early Leave (HD) (BD)** | 0.5 day |
| **Early Departure (BD)** | 0.5 day |
| **Marriage Leave (BD)** | No deduction |
| **Paternity Leave (BD)** | No deduction |
| **Maternity Leave (BD)** | No deduction |
| **Bereavement Leave (BD)** | No deduction |

### Special
| Status | Description | Color |
|--------|-------------|-------|
| Weekend | OFF day | Gray |

---

## Late Penalty Scale (EGP)

```
1st Late  →  100 EGP
2nd Late  →  200 EGP
3rd Late  →  500 EGP + Warning
4th+ Late →  500 EGP each
```

---

## Missing Punch Rules

```
1-3 occurrences  →  No deduction
4+ occurrences   →  0.5 day each
6+ occurrences   →  Warning issued
```

---

## Color Legend

| Color | Meaning |
|-------|---------|
| 🟢 Green | No deduction |
| 🟡 Yellow | Late penalty |
| 🔴 Red | Absent (2 days) |
| 🟣 Pink | Missing punch |
| 🟠 Orange | Half day deduction |
| 🔵 Light Blue | Leave with deduction |
| 🟪 Purple | Backdated leaves (BD) |
| ⬜ Gray | Weekend/OFF day |

---

## How to Change Status

1. Open the exported Excel file
2. Go to **Summary Report** sheet
3. Click on any date cell (columns D onwards)
4. Click the dropdown arrow
5. Select new justification
6. **Penalties sheet updates automatically!**

---

## Tips

- Use **Late (Approved)** instead of **Late** to remove penalty
- Use **Early Departure (Approved)** to remove half-day deduction
- **Missing Punch (Justified)** still counts toward threshold
- Press **Ctrl+Shift+F9** to force recalculate if needed

---

## Backdated Leaves (BD)

### When to Use
- Leave from a previous month not recorded at the time
- Transferring leave records to current payroll period

### How to Add
1. Open Leave Sheet → Go to current month (e.g., Jan)
2. Find employee by CRM
3. Select any date in payroll period (e.g., Jan 1)
4. Choose `Annual Leave (BD)` or other (BD) variant from dropdown

### Example
```
November leave → Record as "Sick Leave (BD)" on Jan 1st
Result: Counts in January penalties, marked purple for HR review
```

### In Penalties Sheet
- Column Y shows "Backdated Leaves" count
- Purple highlighting for easy identification
- HR can review and adjust if needed
