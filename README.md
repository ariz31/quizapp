# General Quiz App

A Google Apps Script quiz web app that can be used for Civil Engineering or any other subject. The repository is optimized for simple copy-paste setup.

**Live Environment:** [Civil Engineering Quiz App](https://www.arizval.com/civil-engineering/applications/civil-engineering-quiz-app)

## Files to copy into Apps Script

Copy only these files into Apps Script:

- `quiz-app/Code.gs` - single backend file with setup, question loading, result logging, live preview, analytics, diagnostics, and prompt templates
- `quiz-app/QuizPage.html` - student quiz UI, feedback modes, timer, review screen, and analytics UI

Optional supporting prompt templates:

- `quiz-app/prompts/general-question-bank-prompts.md`

## What the app does

- Loads questions from a Google Sheet
- Works for Civil Engineering, general education, review classes, board exam preparation, company training, certification review, and other subject areas
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

## Backend improvements in `Code.gs`

- Single backend file for setup, quiz loading, result saving, live preview, analytics, and prompt templates
- Configurable app metadata with `setAppConfig({ title: 'Your Quiz Title', subject: 'Your Subject' })`
- Flexible question header aliases such as `Question`, `Stem`, `Option A`, `Choice A`, `Correct Answer`, and `Rationale`
- Duplicate Question ID detection
- Invalid-row diagnostics with row-level samples
- Safer setup that recognizes equivalent headers before appending missing columns
- Difficulty normalization for common labels such as Basic, Medium, Hard, and Advanced
- Analytics recommendations for weak areas, duplicate IDs, invalid rows, and top missed questions
- Prompt endpoint through `getPromptTemplates()` for subject-agnostic question generation workflows

## Analytics included in `Code.gs`

- Live filter preview through `getQuestionPreview()`
- Dashboard through `getQuizAnalytics()`
- Weak-area ranking
- Score distribution
- Daily trends
- Option-error patterns
- Median score
- Student improvement summaries
- Question-bank stats with valid and invalid row counts
- Duplicate Question ID count
- Privacy-conscious analytics display with masked student identifiers

## Sheet structure

Create a Google Sheet with these tabs. Running `ensureSetup()` will create missing tabs and headers automatically. It will not delete existing rows.

### `Questions`

Recommended headers:

```text
Question ID, Category, Subject, Topic, Difficulty, Question Text, OptionA, OptionB, OptionC, OptionD, ImageURL, Answer, Explanation
```

Supported aliases include:

- `Question`, `QuestionText`, `Stem`, or `Prompt` for `Question Text`
- `Option A`, `Choice A`, `A` for `OptionA`
- `Correct Answer`, `CorrectAnswer`, `Key`, or `Answer Key` for `Answer`
- `Rationale`, `Solution`, or `Feedback` for `Explanation`
- `Course`, `Discipline`, or `Subject Name` for `Subject`

### `Users`

Stores one row per completed quiz attempt.

### `Responses`

Stores one row per answered question.

## Setup

1. Create a new Google Sheet.
2. Open Extensions > Apps Script.
3. Replace the default `Code.gs` with `quiz-app/Code.gs`.
4. Add an HTML file named `QuizPage` and paste `quiz-app/QuizPage.html`.
5. In `Code.gs`, either replace `DEFAULT_SPREADSHEET_ID` with your Sheet ID or run `setSpreadsheetId('YOUR_SHEET_ID')` once.
6. Optional: run `setAppConfig({ title: 'Your Quiz App', subject: 'Your Subject' })`.
7. Run `ensureSetup()` once and approve permissions.
8. Add or import questions into the `Questions` sheet.
9. Deploy as a web app.
10. Open the web app URL to use the quiz and analytics dashboard.

## Generalized prompts

Use `quiz-app/prompts/general-question-bank-prompts.md` to generate, review, calibrate, improve, and reframe question banks for different subjects.

The prompts are designed for this app schema and can be reused for:

- board exam review
- classroom quizzes
- company training
- certification prep
- language learning
- science and math lessons
- humanities and social science review
- technical skills assessment

## Data preservation

- `ensureSetup()` creates missing tabs only when needed.
- Missing headers are appended to the end of the header row only when no equivalent header already exists.
- Existing rows are not cleared, replaced, or reordered.
- New analytics fields are additive, so older rows can stay in place.
- Analytics in `Code.gs` reads existing rows and returns summaries; it does not write to the spreadsheet.

## Important

This repo is designed for a copy-paste Apps Script workflow, not a local Node/clasp workflow.

For production, use a dedicated quiz spreadsheet and restrict web app access if the quiz is intended only for a class or review group.

## License

MIT License. See `LICENSE` for details.
