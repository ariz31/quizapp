# Civil Engineering Quiz App

A Google Apps Script quiz web app for Civil Engineering board exam review. The repository is optimized for simple copy-paste setup by less technical users and now keeps only deployable app files plus project essentials.

**Live Environment:** [Civil Engineering Quiz App](https://www.arizval.com/civil-engineering/applications/civil-engineering-quiz-app)

## Files to copy into Apps Script

You only need two app files:

- `quiz-app/Code.gs` - all backend logic in one file
- `quiz-app/QuizPage.html` - student quiz setup, quiz-taking UI, feedback modes, timer, review screen, and results screen

No extra `.gs` helper files, prompt files, or legacy artifacts are required.

## What the app does

- Loads questions from a Google Sheet
- Lets students filter by Category, Subject, Topic, and Difficulty
- Supports timed or unlimited quiz attempts
- Supports **Instant Feedback** mode and **Exam Mode**
- Randomizes question order and answer choices
- Allows students to mark questions for review
- Tracks time spent per question
- Shows a full answer review after completion
- Lets students retry missed questions locally after a completed attempt
- Saves quiz summaries to the `Users` sheet
- Saves per-question answers to the `Responses` sheet
- Uses `LockService` and batch writes for safer concurrent submissions
- Works on desktop and mobile layouts

## Added improvements

- Hardened Apps Script backend with `ensureSetup()`
- Optional `setSpreadsheetId('YOUR_SHEET_ID')` helper for non-technical setup
- Required sheet/header creation without deleting existing data
- Difficulty-aware question filtering
- Question-bank stats endpoint with valid/invalid row counts
- Safer result payload validation before saving
- Safer frontend rendering using `textContent` instead of injecting question text with `innerHTML`
- Mobile-first card layout with accessible buttons, focus states, status panel, and toast messages
- Feedback Mode selector: instant feedback or end-of-quiz explanations
- Results review list showing correctness, selected answer, correct answer, explanation, review mark, and time spent
- Retry Missed Questions flow for focused practice
- Local browser memory for student name and ID number
- Keyboard shortcuts: press `A`, `B`, `C`, or `D` to answer; press `Enter` to continue after answering
- Expanded response logs for analytics: question number, marked-for-review flag, time spent, quiz duration, accuracy band, and feedback mode
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

Stores one row per completed quiz attempt. Newer installs include started/submitted time, quiz duration, score, percentage, accuracy band, selected filters, feedback mode, and quiz mode JSON.

### `Responses`

Stores one row per answered question. Newer installs include question number, selected answer, correct answer, correctness, timeout status, marked-for-review status, time spent, and metadata.

## Setup

1. Create a new Google Sheet.
2. Open **Extensions > Apps Script**.
3. Replace the default `Code.gs` with `quiz-app/Code.gs`.
4. Add an HTML file named `QuizPage` and paste `quiz-app/QuizPage.html`.
5. In `Code.gs`, either replace `DEFAULT_SPREADSHEET_ID` with your Sheet ID or run `setSpreadsheetId('YOUR_SHEET_ID')` once.
6. Run `ensureSetup()` once and approve permissions.
7. Add or import questions into the `Questions` sheet.
8. Deploy as a web app.
9. Open the web app URL to use the quiz.

## Question import workflow

Prepare your question bank using the `Questions` sheet headers. Rows may be typed directly, pasted from a spreadsheet, or imported from a CSV using this column order:

```text
Question ID;Category;Subject;Topic;Difficulty;Question Text;OptionA;OptionB;OptionC;OptionD;ImageURL;Answer;Explanation
```

After importing, run the web app setup screen and verify the valid/invalid question counts before allowing students to take the quiz.

## Data preservation

This repository is designed to preserve existing spreadsheet data:

- `ensureSetup()` creates missing tabs only when needed.
- Missing headers are appended to the end of the header row.
- Existing rows are not cleared, replaced, or reordered.
- New analytics fields are additive, so older rows can stay in place.

## Important

This repo is designed for a copy-paste Apps Script workflow, not a local Node/clasp workflow. Keep all backend logic inside `Code.gs` unless the project is intentionally migrated later.

For production, avoid exposing private spreadsheet data. Use a dedicated quiz spreadsheet and restrict web app access if the quiz is intended only for a class or review group.

## License

MIT License. See `LICENSE` for details.
