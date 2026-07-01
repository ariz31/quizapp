# Civil Engineering Quiz App

A Google Apps Script quiz web app for Civil Engineering board exam review. The repository is optimized for simple copy-paste setup and keeps only deployable app files plus project essentials.

**Live Environment:** [Civil Engineering Quiz App](https://www.arizval.com/civil-engineering/applications/civil-engineering-quiz-app)

## Files to copy into Apps Script

Copy these app files into Apps Script:

- `quiz-app/Code.gs`
- `quiz-app/QuizPage.html`
- `quiz-app/Analytics.gs`
- `quiz-app/AnalyticsAdvanced.gs`
- `quiz-app/ZZ_AnalyticsDashboard.gs`

No legacy prompt files or documentation-only folders are required.

## What the app does

- Loads questions from a Google Sheet
- Lets students filter by Category, Subject, Topic, and Difficulty
- Shows a live matching-question count before the quiz starts
- Supports timed or unlimited quiz attempts
- Supports Instant Feedback mode and Exam Mode
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

## Analytics improvements

- Live filter preview through `getQuestionPreview()`
- Dashboard through `getQuizAnalytics()`
- Weak-area ranking
- Score distribution
- Daily trends
- Option-error patterns
- Median score
- Student improvement summaries
- Question-bank stats with valid and invalid row counts
- Privacy-conscious analytics display with masked student identifiers

## Sheet structure

Create a Google Sheet with these tabs. Running `ensureSetup()` will create missing tabs and headers automatically. It will not delete existing rows.

### `Questions`

Required headers:

```text
Question ID, Category, Subject, Topic, Difficulty, Question Text, OptionA, OptionB, OptionC, OptionD, ImageURL, Answer, Explanation
```

### `Users`

Stores one row per completed quiz attempt.

### `Responses`

Stores one row per answered question.

## Setup

1. Create a new Google Sheet.
2. Open Extensions > Apps Script.
3. Replace the default `Code.gs` with `quiz-app/Code.gs`.
4. Add script files for `Analytics`, `AnalyticsAdvanced`, and `ZZ_AnalyticsDashboard`, then paste the matching files.
5. Add an HTML file named `QuizPage` and paste `quiz-app/QuizPage.html`.
6. In `Code.gs`, either replace `DEFAULT_SPREADSHEET_ID` with your Sheet ID or run `setSpreadsheetId('YOUR_SHEET_ID')` once.
7. Run `ensureSetup()` once and approve permissions.
8. Add or import questions into the `Questions` sheet.
9. Deploy as a web app.
10. Open the web app URL to use the quiz and analytics dashboard.

## Data preservation

- `ensureSetup()` creates missing tabs only when needed.
- Missing headers are appended to the end of the header row.
- Existing rows are not cleared, replaced, or reordered.
- New analytics fields are additive, so older rows can stay in place.
- Analytics files read existing rows and return summaries; they do not write to the spreadsheet.

## Important

This repo is designed for a copy-paste Apps Script workflow, not a local Node/clasp workflow.

For production, use a dedicated quiz spreadsheet and restrict web app access if the quiz is intended only for a class or review group.

## License

MIT License. See `LICENSE` for details.
