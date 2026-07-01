# Civil Engineering Quiz App

A Google Apps Script quiz web app for Civil Engineering board exam review. This version is optimized for simple copy-paste setup by less technical users, following the same repository style as the Peer Evaluation Google Apps Script app.

**Live Environment:** [Civil Engineering Quiz App](https://www.arizval.com/civil-engineering/applications/civil-engineering-quiz-app)

## Files to copy into Apps Script

You only need two app files:

- `quiz-app/Code.gs` - all backend logic in one file
- `quiz-app/QuizPage.html` - student quiz setup, quiz-taking UI, feedback, timer, and results screen

Optional supporting file:

- `quiz-app/ai prompts/ce quiz maker.txt` - prompt/instructions for generating spreadsheet-ready Civil Engineering quiz questions

No extra `.gs` helper files are required.

## What the app does

- Loads questions from a Google Sheet
- Lets students filter by Category, Subject, Topic, and Difficulty
- Supports timed or unlimited quiz attempts
- Randomizes question order and answer choices
- Shows instant feedback and explanations
- Saves quiz summaries to the `Users` sheet
- Saves per-question answers to the `Responses` sheet
- Uses `LockService` and batch writes for safer concurrent submissions
- Works on desktop and mobile layouts

## Added improvements

- Hardened Apps Script backend with `ensureSetup()`
- Optional `setSpreadsheetId('YOUR_SHEET_ID')` helper for non-technical setup
- Required sheet/header creation without deleting existing data
- Difficulty-aware question filtering
- Safer result payload validation before saving
- Safer frontend rendering using `textContent` instead of injecting question text with `innerHTML`
- Mobile-first card layout with accessible buttons, focus states, status panel, and toast messages
- Clear save status on the results screen
- Root `.gitignore` and MIT `LICENSE`

## Sheet structure

Create a Google Sheet with these tabs. Running `ensureSetup()` will create missing tabs and headers automatically.

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

Stores one row per completed quiz attempt, including name, ID number, total score, percentage, selected filters, and quiz mode.

### `Responses`

Stores one row per answered question, including selected answer, correct answer, correctness, timeout status, and metadata.

## Setup

1. Create a new Google Sheet.
2. Open **Extensions > Apps Script**.
3. Replace the default `Code.gs` with `quiz-app/Code.gs`.
4. Add an HTML file named `QuizPage` and paste `quiz-app/QuizPage.html`.
5. In `Code.gs`, either:
   - replace `DEFAULT_SPREADSHEET_ID` with your Sheet ID, or
   - run `setSpreadsheetId('YOUR_SHEET_ID')` once from the Apps Script editor.
6. Run `ensureSetup()` once and approve permissions.
7. Add or import questions into the `Questions` sheet.
8. Deploy as a web app.
9. Open the web app URL to use the quiz.

## Deployment settings

Recommended Google Apps Script deployment settings:

- **Type:** Web app
- **Execute as:** Me
- **Who has access:** Anyone, or your selected school/user group

## Question generation workflow

Use `quiz-app/ai prompts/ce quiz maker.txt` to generate rows for the `Questions` sheet. The expected output format is semicolon-delimited and should match this column order:

```text
Question ID;Category;Subject;Topic;Difficulty;Question Text;OptionA;OptionB;OptionC;OptionD;ImageURL;Answer;Explanation
```

After generating questions, paste/import them into the `Questions` sheet using the same headers.

## Important

This repo is designed for a copy-paste Apps Script workflow, not a local Node/clasp workflow. Keep all backend logic inside `Code.gs` unless the project is intentionally migrated later.

For production, avoid exposing private spreadsheet data. Use a dedicated quiz spreadsheet and restrict web app access if the quiz is intended only for a class or review group.

## License

MIT License. See `LICENSE` for details.
