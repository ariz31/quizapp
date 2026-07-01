# General Question Bank Prompt Templates

Use these prompts for any subject, not only Civil Engineering. Replace bracketed placeholders before use.

## 1. Full Question Bank Generator

You are an expert assessment designer for [SUBJECT]. Create a question bank for [LEARNER LEVEL] learners.

Goal: Generate high-quality multiple-choice questions that test understanding, application, and problem solving.

Output format: CSV-compatible rows using these columns only:

```text
Question ID;Category;Subject;Topic;Difficulty;Question Text;OptionA;OptionB;OptionC;OptionD;ImageURL;Answer;Explanation
```

Rules:

- Generate [NUMBER] questions.
- Category should be a broad learning area.
- Subject should be the course or domain name.
- Topic should be the specific lesson or skill.
- Difficulty must be Easy, Normal, or Difficult unless I provide another scale.
- Answer must be A, B, C, or D.
- Exactly one option must be correct.
- Distractors must be plausible and based on common mistakes.
- Explanation must teach the concept, not only state the answer.
- Do not include markdown tables.
- Use semicolons as separators.
- Leave ImageURL blank unless I provide image URLs.

Subject: [SUBJECT]
Learner level: [LEARNER LEVEL]
Topics to cover: [TOPICS]
Question count: [NUMBER]
Special constraints: [CONSTRAINTS]

## 2. Question Quality Reviewer

Review the following question bank for quality, clarity, answer-key correctness, duplicate questions, weak distractors, ambiguous wording, and topic coverage.

Return:

1. Critical issues that must be fixed.
2. Suggested improvements.
3. Rows that may have wrong answers.
4. Duplicate or near-duplicate questions.
5. Coverage gaps by topic and difficulty.
6. A corrected CSV-ready version only for rows that need changes.

Question bank:

[PASTE QUESTIONS]

## 3. Difficulty Calibration Prompt

Calibrate the difficulty labels for this question bank.

Use this standard:

- Easy: direct recall, basic recognition, one-step application.
- Normal: multi-step reasoning, moderate application, common exam-level problem.
- Difficult: advanced reasoning, traps, synthesis, multi-concept problem, or high cognitive load.

Return a list of rows whose difficulty should change and explain why.

Question bank:

[PASTE QUESTIONS]

## 4. Explanation Improvement Prompt

Improve the explanations in this question bank so learners understand why the correct answer is correct and why the most tempting wrong options are wrong.

Keep the same columns and answers. Do not change the correct answer unless it is clearly wrong.

Question bank:

[PASTE QUESTIONS]

## 5. Topic Expansion Prompt

Create additional questions that match the style and schema of my existing question bank.

Existing sample:

[PASTE SAMPLE QUESTIONS]

Add questions for these topics:

[TOPICS]

Use the same output columns:

```text
Question ID;Category;Subject;Topic;Difficulty;Question Text;OptionA;OptionB;OptionC;OptionD;ImageURL;Answer;Explanation
```

## 6. Subject Reframing Prompt

Convert this question-generation prompt or question bank so it works for [NEW SUBJECT] while keeping the same schema and app compatibility.

Requirements:

- Keep the same CSV columns.
- Make category, subject, topic, and difficulty labels appropriate for [NEW SUBJECT].
- Keep explanations instructional.
- Avoid subject-specific assumptions from the original domain.

Input:

[PASTE ORIGINAL PROMPT OR QUESTIONS]
