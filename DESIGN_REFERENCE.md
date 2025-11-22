# 509 Dashboard - Design Reference Guide

Based on visual references from preferred dashboard examples

---

## 🎨 Color Scheme Preferences

### Primary Color Palettes

**Option 1: Professional Green/Teal Theme**
- Primary: `#5A9E8E` (Sage Green/Teal)
- Secondary: `#F4C430` (Golden Yellow)
- Accent: `#FFFFFF` (White backgrounds)
- Headers: `#4A7C6D` (Deep Teal)
- Text: `#333333` (Dark Gray)

**Option 2: Modern Purple/Blue Theme**
- Primary: `#7B68EE` (Medium Purple)
- Secondary: `#00C9FF` (Bright Cyan)
- Accent: `#FF6B9D` (Pink)
- Headers: `#4A5568` (Slate Gray)
- Background: `#F7FAFC` (Light Gray)

**Option 3: Dark/Modern Theme**
- Primary: `#1E293B` (Dark Blue-Gray)
- Secondary: `#0EA5E9` (Bright Blue)
- Accent: `#10B981` (Green)
- Warning: `#F59E0B` (Orange)
- Error: `#EF4444` (Red)

### Status Colors (Universal)
- **Success/Positive:** `#10B981` (Green)
- **Warning/Due Soon:** `#F59E0B` (Orange/Yellow)
- **Error/Overdue:** `#EF4444` (Red)
- **Info/Neutral:** `#3B82F6` (Blue)
- **In Progress:** `#8B5CF6` (Purple)

---

## 📊 KPI Card Design

### Layout Structure
```
┌─────────────────────────────┐
│   KPI LABEL                 │
│                             │
│        1,234                │  ← Large, bold number
│                             │
│   ↑ 12.5% vs last month    │  ← Trend indicator
└─────────────────────────────┘
```

### Key Features
- **Card Style:** Light background with subtle border or shadow
- **Number Size:** 32-48pt, bold weight
- **Label:** 12-14pt, uppercase or title case
- **Trend Indicators:**
  - Green up arrow (▲) for positive
  - Red down arrow (▼) for negative
  - Percentage in matching color
- **Spacing:** Generous padding (15-20px)
- **Alignment:** Center or left-aligned based on context

### Examples from References
1. **Gaming KPI Dashboard:** Green/purple header, white cards, up/down triangles
2. **CRM Dashboard:** Clean white cards, colored text for metrics
3. **Warehouse Dashboard:** Purple header bar, boxed metrics with labels

---

## 📈 Chart Preferences

### 1. Donut Charts
**Best For:** Status breakdowns, category distribution
- **Colors:** Use 3-5 distinct colors maximum
- **Center:** Display total count or primary metric
- **Labels:** Show percentages and category names
- **Examples Seen:**
  - Grievances by Status (Filed, Resolved, In Mediation)
  - Task completion (Not Started, In Progress, Complete)
  - Member distribution by Unit

### 2. Horizontal Bar Charts
**Best For:** Top 10 lists, rankings, comparisons
- **Color:** Single color with gradient OR category-based colors
- **Labels:** Clear left-aligned labels
- **Values:** Display at end of bars
- **Examples Seen:**
  - Top 10 Grievance Types
  - Top 10 Locations
  - Performance rankings

### 3. Vertical Column Charts
**Best For:** Time series, step comparisons, counts
- **Color:** Single color for single series, multiple for comparisons
- **Spacing:** Adequate gap between columns
- **Grid Lines:** Subtle horizontal lines
- **Examples Seen:**
  - Grievances by Step
  - Monthly trends
  - Cost comparisons by category

### 4. Line Charts
**Best For:** Trends over time, performance tracking
- **Line Weight:** 2-3px thickness
- **Points:** Optional markers at data points
- **Area Fill:** Optional subtle fill under line
- **Examples Seen:**
  - All deals over time (cumulative)
  - Distance/pace over time
  - Monthly filing trends

### 5. Stacked Bar Charts
**Best For:** Multi-category comparisons, workload distribution
- **Colors:** Distinct colors per category (max 5)
- **Legend:** Clear legend with color coding
- **Orientation:** Horizontal or vertical based on labels
- **Examples Seen:**
  - By status per owner
  - Purchase by week across units
  - Inventory stock levels

---

## 🎯 Dashboard Layout Principles

### Grid Structure
```
┌─────────────────────────────────────────────────────────┐
│                    HEADER / TITLE                       │
├──────────────┬──────────────┬──────────────┬────────────┤
│   KPI Card   │   KPI Card   │   KPI Card   │  KPI Card  │  ← Row 1: Key Metrics
├──────────────┴──────────────┴──────────────┴────────────┤
│                                                          │
│              Status Overview (Donut Chart)               │  ← Row 2: Primary Visual
│                                                          │
├────────────────────────────┬─────────────────────────────┤
│                            │                             │
│   Chart 2                  │   Chart 3                   │  ← Row 3: Secondary Charts
│   (Bar Chart)              │   (Column Chart)            │
│                            │                             │
├────────────────────────────┴─────────────────────────────┤
│                                                          │
│              Data Table (Top 10 / Recent Items)          │  ← Row 4: Detailed Data
│                                                          │
└──────────────────────────────────────────────────────────┘
```

### Spacing Guidelines
- **Section Padding:** 20-30px between major sections
- **Card Spacing:** 15-20px gap between cards
- **Chart Margins:** 10-15px internal padding
- **Mobile Responsive:** Stack columns on smaller screens

### Header Design
- **Background:** Solid color bar (green, purple, blue, or dark)
- **Text:** White or contrasting color
- **Height:** 60-80px
- **Font:** 24-32pt, bold
- **Elements:** Title + optional date range/filters

---

## 📋 Data Table Styling

### Table Features from References

**Header Row:**
- Bold text, colored background (matching theme)
- All caps or title case
- Centered or left-aligned

**Data Rows:**
- Alternating row colors (white / light gray)
- Hover effect for interactivity
- Clear borders or spacing

**Conditional Formatting:**
- Color-coded cells based on values
- Red/yellow/green indicators
- Progress bars within cells
- Up/down arrows for trends

**Example from Gaming KPI:**
- Green header row
- Purple and green sections for MTD/YTD
- Triangle indicators (▲▼) for performance
- Percentage comparisons with color coding

---

## 🎭 Theme Recommendations for 509 Dashboard

### Recommended: **Professional Union Theme**

**Primary Colors:**
- Header: `#2563EB` (Union Blue) or `#DC2626` (Solidarity Red)
- Background: `#F9FAFB` (Light Gray)
- Cards: `#FFFFFF` (White)
- Accent: `#059669` (Green for wins)

**KPI Cards:**
- White background
- Subtle shadow: `0 1px 3px rgba(0,0,0,0.1)`
- Border: `1px solid #E5E7EB`
- Hover effect: slight elevation

**Charts:**
1. **Grievances by Status** - Donut chart with:
   - Filed: `#3B82F6` (Blue)
   - In Progress: `#8B5CF6` (Purple)
   - Resolved: `#10B981` (Green)
   - In Mediation: `#F59E0B` (Orange)
   - In Arbitration: `#EF4444` (Red)

2. **Grievances by Step** - Column chart with:
   - Step I: `#60A5FA` (Light Blue)
   - Step II: `#3B82F6` (Blue)
   - Step III: `#1E40AF` (Dark Blue)

3. **Top 10 Grievance Types** - Horizontal bars:
   - Single color: `#3B82F6` (Blue)
   - Gradient from dark to light

4. **Win Rate Outcomes** - Donut chart:
   - Won: `#10B981` (Green)
   - Lost: `#EF4444` (Red)
   - Settled: `#F59E0B` (Orange)
   - Withdrawn: `#6B7280` (Gray)

**Status Indicators:**
- Overdue: Red background `#FEE2E2`, red text `#DC2626`
- Due Soon: Yellow background `#FEF3C7`, yellow text `#D97706`
- On Track: Green background `#D1FAE5`, green text `#059669`

---

## 💡 Key Design Insights from References

### What Works Well:

1. **Clean Hierarchy**
   - Most important metrics at the top
   - Visual interest in the middle
   - Detailed data at the bottom

2. **Color Psychology**
   - Green = Success, positive outcomes
   - Red = Urgent, negative outcomes
   - Blue = Neutral, informational
   - Purple = In progress, pending
   - Orange = Warning, attention needed

3. **Visual Balance**
   - Mix of numbers (KPIs) and visuals (charts)
   - Don't overcrowd - white space is good
   - 2-3 column layouts work best

4. **Professional Polish**
   - Consistent fonts throughout
   - Aligned elements
   - Subtle shadows and borders
   - Smooth color transitions

5. **Data Storytelling**
   - Each chart answers a question
   - KPIs provide quick overview
   - Tables offer detailed drill-down
   - Trend indicators show movement

---

## 🔧 Implementation Notes for Google Sheets

### Chart Styling Settings
- **Chart Border:** None or subtle 1px
- **Background:** Transparent or white
- **Legend Position:** Right or bottom
- **Gridlines:** Light gray, horizontal only
- **Font:** Arial or Roboto, 10-12pt

### Conditional Formatting Rules
```
Overdue (30+ days):    Background: #FEE2E2, Text: #991B1B
Overdue (14-30 days):  Background: #FEF3C7, Text: #92400E
Overdue (< 14 days):   Background: #FEF9C3, Text: #854D0E
On Track:              Background: #D1FAE5, Text: #065F46
```

### Number Formatting
- **Large Numbers:** 1,234 (comma separators)
- **Percentages:** 45.2% (one decimal)
- **Currency:** $1,234.56
- **Dates:** MM/DD/YYYY or Nov 22, 2025

---

## 📸 Reference Images

The following dashboard examples were analyzed:

1. **CRM Dashboard** - Clean white KPI cards with colored metrics
2. **Gaming KPI Dashboard** - Green/purple theme with triangle indicators
3. **Google Sheets Dark Dashboard** - Dark theme with bright blue/green accents
4. **Light Warehouse Dashboard** - Purple header, pastel visualizations
5. **Running Analytics** - Green/yellow theme with large metric displays
6. **Portfolio Dashboard** - Clean light mode with donut and bar charts
7. **Employee Performance Dashboard** - Multi-colored donut charts with data table

All images located in: `/home/user/delete/`

---

**Design Philosophy:** Professional, data-focused, union-friendly, and accessible to all skill levels.
