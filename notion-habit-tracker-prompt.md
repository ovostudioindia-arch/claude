# Notion AI Prompt — Habit Tracker Template

Paste the following into Notion AI:

---

Create a complete, professional-grade Habit Tracker system on this single page. This should function as a sellable Notion template — visually polished, beginner-friendly, and powerful enough for advanced users. Follow these instructions precisely:

---

## PAGE STRUCTURE

At the top of the page, create a **header section** with:

- Title: **"Habit Tracker"** (large heading)
- A short italic subtitle: *"Small daily improvements lead to stunning results."*
- A horizontal divider

Below the header, create a **Dashboard Stats** callout block (💡 icon) with a single row of inline stats using formulas (these pull from the database below). Display them on one line separated by `  |  `:

```
🔥 Current Streak: 12 days  |  🏆 Best Streak: 30 days  |  ✅ Today: 5/8 done  |  📊 This Week: 78%  |  ▓▓▓▓▓▓▓░░░ 78%
```

The progress bar (▓ and ░) should visually represent the weekly completion percentage. Use 10 blocks total.

Below the dashboard, add a **Quick Links** toggle block containing:
- "How to Use This Tracker" — a short 4-5 bullet guide explaining how to add habits, log completions, and read the stats.
- "Customize Your Categories" — explains how to edit the category property to add personal groupings.

---

## DATABASE: Habit Log

Create **one main database** called **"Habit Log"** with the following properties:

| Property | Type | Details |
|---|---|---|
| **Habit** | Title | Name of the habit (e.g., "Meditate", "Read 30 min") |
| **Date** | Date | The date this entry is for |
| **Done** | Checkbox | Whether the habit was completed |
| **Category** | Select | Options: 🏋️ Health, 🧠 Mind, 💼 Work, 💰 Finance, 🎨 Creative, 🤝 Social |
| **Frequency** | Select | Options: Daily, Weekdays, Weekends, 3x/Week, Weekly |
| **Priority** | Select | Options: 🔴 High, 🟡 Medium, 🟢 Low |
| **Mood After** | Select | Options: 😊 Great, 🙂 Good, 😐 Neutral, 😔 Low |
| **Notes** | Text | Optional reflection or quick note |
| **Streak** | Formula | Calculates consecutive days completed for this habit |
| **Week** | Formula | Extracts the ISO week number from the Date property |
| **Month** | Formula | Extracts the month name from the Date property (e.g., "April") |
| **Completion** | Formula | Returns "✅" if Done is checked, "⬜" if not |
| **Progress Label** | Formula | Generates a visual like: `▓▓▓▓░░░░░░ 40%` based on that habit's weekly completion rate |

---

## DATABASE VIEWS (all within the same database)

Create the following **6 views** of the Habit Log database, each with specific filters, sorts, and layouts:

### View 1: "📅 Today" (Table View)
- **Filter:** Date = Today
- **Sort:** Category (ascending), then Priority (🔴 first)
- **Visible columns:** Habit, Done, Category, Priority, Mood After, Notes
- **Purpose:** Daily check-in. Users open this, tick off habits, optionally log mood and notes.

### View 2: "📆 This Week" (Calendar View)
- **Filter:** Date is within the current week
- **Layout:** Calendar, grouped by date
- **Card preview:** Show Habit name and Completion (✅/⬜)
- **Purpose:** Weekly bird's-eye view of what's done and what's missed.

### View 3: "📊 Weekly Summary" (Board/Kanban View)
- **Group by:** Category
- **Filter:** Date is within the current week
- **Card properties:** Habit, Completion, Progress Label
- **Purpose:** See how each life area is performing this week. Each column is a category (Health, Mind, Work, etc.) with habit cards showing progress bars.

### View 4: "📈 Streaks & Stats" (Table View)
- **Filter:** Done = checked
- **Sort:** Streak (descending)
- **Visible columns:** Habit, Streak, Category, Progress Label, Date
- **Purpose:** Leaderboard of your best-performing habits. Motivation view.

### View 5: "🗂️ All Habits" (Table View)
- **Filter:** None (show everything)
- **Sort:** Date (descending)
- **Visible columns:** All properties
- **Purpose:** Master log. Full data access for power users who want to filter, export, or audit.

### View 6: "🎯 By Category" (Gallery View)
- **Group by:** Category
- **Card preview:** Habit name, Completion, Mood After, Progress Label
- **Card size:** Small
- **Purpose:** Visual, aesthetic overview grouped by life area. Great for scanning balance across categories.

---

## VISUAL & UX GUIDELINES

- Use **callout blocks** with emojis for the dashboard and quick-links sections (not plain text).
- Use **toggles** to keep instructional content collapsed and the page clean.
- Place the **"📅 Today"** view as the **first/default view** — this is what users interact with daily.
- Add a **divider** between the dashboard section and the database.
- Use **color-coded selects** for Category and Priority so the board and gallery views are visually distinct.
- Keep the page to **one database only** — all views derive from the same Habit Log.
- Make sure the progress bar formula (`▓░`) works by calculating: (completed entries for this habit this week) / (total entries for this habit this week) × 100, then mapping to 10-character bar.

---

## FORMULA GUIDANCE

For the **Progress Label** formula, use logic like:

```
lets(
  pct, <calculate weekly completion percentage>,
  filled, floor(pct / 10),
  empty, 10 - filled,
  "▓".repeat(filled) + "░".repeat(empty) + " " + format(round(pct)) + "%"
)
```

For the **Streak** formula, calculate consecutive checked days backward from today for the same Habit name.

For **Week**, use: `formatDate(prop("Date"), "W")`

For **Month**, use: `formatDate(prop("Date"), "MMMM")`

---

## SAMPLE DATA

Pre-populate the database with **7 days of sample data** (past week) for these habits:

1. 🧘 Meditate (10 min) — Category: 🧠 Mind — Frequency: Daily — Priority: 🔴 High
2. 💧 Drink 2L Water — Category: 🏋️ Health — Frequency: Daily — Priority: 🔴 High
3. 📖 Read 30 min — Category: 🧠 Mind — Frequency: Daily — Priority: 🟡 Medium
4. 🏃 Exercise — Category: 🏋️ Health — Frequency: 3x/Week — Priority: 🔴 High
5. 📝 Journal — Category: 🎨 Creative — Frequency: Daily — Priority: 🟡 Medium
6. 💰 Review Budget — Category: 💰 Finance — Frequency: Weekly — Priority: 🟢 Low
7. 📵 No Phone Before 9am — Category: 🧠 Mind — Frequency: Weekdays — Priority: 🟡 Medium
8. 🤝 Connect with a Friend — Category: 🤝 Social — Frequency: 3x/Week — Priority: 🟢 Low

Mix completions realistically — aim for ~70-80% overall completion, with some habits having perfect streaks and others with gaps. Vary the "Mood After" entries. Leave a few "Notes" entries filled in with short reflections like "Felt really focused after this" or "Skipped — was too tired".

---

Generate this entire system now on this page. Prioritize functionality and visual clarity. Everything should be on one page with one database and multiple views.
