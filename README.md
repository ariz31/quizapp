# Civil Engineering Quiz App

A Google Apps Script quiz web app for Civil Engineering board exam review. The repository is optimized for simple copy-paste setup and keeps only deployable app files plus project essentials.

**Live Environment:** [Civil Engineering Quiz App](https://www.arizval.com/civil-engineering/applications/civil-engineering-quiz-app)

## Files to copy into Apps Script

Copy these app files into Apps Script:

- `quiz-app/Code.gs` - core quiz backend, sheet setup, question loading, and result saving
- `quiz-app/QuizPage.html` - student quiz UI, feedback modes, timer, review screen, and analytics UI
- `quiz-app/Analytics.gs` - base faculty analytics helpers and live question-preview helpers
- `quiz-app/AnalyticsAdvanced.gs` - deeper read-only analytics calculations
- `quiz-app/ZZ_AnalyticsDashboard.gs` - dashboard endpoint for the richer analytics model

No legacy prompt files or documentation-only folders are required.

## What the app does

- Loads questions from a Google Sheet
- Lets students filter by Category, Subject, Topic, and Difficulty
- Shows a live matching-question count before the quiz starts
- Supports timed or unlimited quiz attempts
- Supports **Instant Feedback** mode and **Exam Mode**
- Randomizes question order and answer choices
- Allows students to mark questions for review
- Tracks time spent per question
- Shows a full answer review after completion
- Lets students retry missed questions locally after a completed attempt
- Saves quiz summaries to the `Users` sheet
- Saves per-question answers to the `Responses` sheet
- Provides a faculty analytics dashboard for attempts, accuracy, weak areas, score distribution, missed questions, recent attempts, and question-bank health
- Uses `LockService` and batch writes for safer concurrent submissions
- Works on desktop and mobile layouts

## Current improvements

- Hardened Apps Script backend with `ensureSetup()`
- Optional `setSpreadsheetId('YOUR_SHEET_ID')` helper for non-technical setup
- Required sheet/header creation without deleting existing data
- Difficulty-aware question filtering
- Live filter preview through `getQuestionPreview()`
- Faculty analytics dashboard through `getQuizAnalytics()`
- Advanced analytics: weak-area ranking, score distribution, daily trends, option-error patterns, median score, and student improvement summaries
- Question-bank stats with valid/invalid row counts
- Safer result payload validation before saving
- Safer frontend rendering using `textContent`
- Mobile-first card layout with accessible buttons, focus states, status panel, and toast messages
- Feedback Mode selector: instant feedback or end-of-quiz explanations
- Results review list showing correctness, selected answer, correct answer, explanation, review mark, and time spent
- Retry Missed Questions flow for focused practice
- Local browser memory for student name and ID number
- Keyboard shortcuts: press `A`, `B`, `C`, or `D` to answer; press `Enter` to continue after answering
- Expanded response logs for analytics: question number, marked-for-review flag, time spent, quiz duration, accuracy band, and feedback mode
- Privacy-conscious analytics display with masked student identifiers
- Legacy prompt artifact removed
- Root `.gitignore` and MIT `LICENSE`

## Sheet structure

Create a Google Sheet with these tabs. Running `ensureSetup()` will create missing tabs and headers automatically. It will not delete existing rows.

### `Questions`

Required headers:

```text
Question ID, Category, Subject, Topic, Difficulty, Question Text, OptionA, OptionB, OptionC, OptionD, ImageURL, Answer, Explanation
```

Rules:

- `Answer` must be `A`, `B`, `C`, or `D`.
- `Difficulty` may be `Easy`, `Normal`, `Difficult`, or any consistent label you prefer.
- `ImageURL` may be blank.
- Keep questions and explanations as plain text. Unicode math symbols are supported.

### `Users`

Stores one row per completed quiz attempt, including started/submitted time, quiz duration, score, percentage, accuracy band, selected filters, feedback mode, and quiz mode JSON.

### `Responses`

Stores one row per answered question, including question number, selected answer, correct answer, correctness, timeout status, marked-for-review status, time spent, and metadata.

## Setup

1. Create a new Google Sheet.
2. Open **Extensions > Apps Script**.
3. Replace the default `Code.gs` with `quiz-app/Code.gs`.
4. Add script files for `Analytics`, `AnalyticsAdvanced`, and `ZZ_AnalyticsDashboard`, then paste the matching `.gs` files.
5. Add an HTML file named `QuizPage` and paste `quiz-app/QuizPage.html`.
6. In `Code.gs`, either replace `DEFAULT_SPREADSHEET_ID` with your Sheet ID or run `setSpreadsheetId('YOUR_SHEET_ID')` once.
7. Run `ensureSetup()` once and approve permissions.
8. Add or import questions into the `Questions` sheet.
9. Deploy as a web app.
10. Open the web app URL to use the quiz and analytics dashboard.

## Analytics dashboard

The faculty analytics dashboard summarizes existing `Users`, `Responses`, and `Questions` data. It does not modify stored rows.

It includes KPI cards, accuracy breakdowns, weak-area ranking, score distribution, daily trend data, top missed questions, masked student summaries, recent attempts, and question-bank health metrics.

For private classes, restrict the web app deployment to your intended users or school domain.

## Data preservation

- `ensureSetup()` creates missing tabs only when needed.
- Missing headers are appended to the end of the header row.
- Existing rows are not cleared, replaced, or reordered.
- New analytics fields are additive, so older rows can stay in place.
- Analytics files read existing rows and return summaries; they do not write to the spreadsheet.

## Important

This repo is designed for a copy-paste Apps Script workflow, not a local Node/clasp workflow. Keep the deployment simple unless the project is intentionally migrated later.

For production, avoid exposing private spreadsheet data. Use a dedicated quiz spreadsheet and restrict web app access if the quiz is intended only for a class or review group.

## License

MIT License. See `LICENSE` for details.
