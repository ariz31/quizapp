/**
 * General Quiz App - Single File Backend
 *
 * Copy this single file into Apps Script as Code.gs. It contains setup,
 * question loading, quiz result logging, live question preview, faculty
 * analytics, and generalized prompt templates. QuizPage.html is the only
 * other required Apps Script file.
 *
 * @version 5.1.0
 * @license MIT
 */

const APP_TITLE = 'Civil Engineering Quiz App';
const DEFAULT_SPREADSHEET_ID = '1qQw7B6sRrTkbGPViBmYqqaJEqwQr-7P0jeNuPirrgpY';
const SCRIPT_PROP_SPREADSHEET_ID = 'SPREADSHEET_ID';
const SCRIPT_PROP_APP_TITLE = 'APP_TITLE';
const SCRIPT_PROP_APP_SUBJECT = 'APP_SUBJECT';

const QUESTIONS_SHEET_NAME = 'Questions';
const RESPONSES_SHEET_NAME = 'Responses';
const USERS_SHEET_NAME = 'Users';

const ALL_VALUE = 'All';
const MAX_QUESTIONS_PER_RUN = 100;
const MAX_DIAGNOSTIC_ROWS = 30;

const Q_COL_ID = 'Question ID';
const Q_COL_CATEGORY = 'Category';
const Q_COL_SUBJECT = 'Subject';
const Q_COL_TOPIC = 'Topic';
const Q_COL_DIFFICULTY = 'Difficulty';
const Q_COL_TEXT = 'Question Text';
const Q_COL_OPTION_A = 'OptionA';
const Q_COL_OPTION_B = 'OptionB';
const Q_COL_OPTION_C = 'OptionC';
const Q_COL_OPTION_D = 'OptionD';
const Q_COL_IMAGE_URL = 'ImageURL';
const Q_COL_ANSWER = 'Answer';
const Q_COL_EXPLANATION = 'Explanation';

const QUESTION_HEADERS = [
  Q_COL_ID,
  Q_COL_CATEGORY,
  Q_COL_SUBJECT,
  Q_COL_TOPIC,
  Q_COL_DIFFICULTY,
  Q_COL_TEXT,
  Q_COL_OPTION_A,
  Q_COL_OPTION_B,
  Q_COL_OPTION_C,
  Q_COL_OPTION_D,
  Q_COL_IMAGE_URL,
  Q_COL_ANSWER,
  Q_COL_EXPLANATION,
];

const HEADER_ALIASES = {};
HEADER_ALIASES[Q_COL_ID] = ['QuestionID', 'Question Id', 'ID', 'Item ID', 'ItemID'];
HEADER_ALIASES[Q_COL_CATEGORY] = ['Area', 'Domain', 'Strand', 'Unit'];
HEADER_ALIASES[Q_COL_SUBJECT] = ['Course', 'Discipline', 'Subject Name'];
HEADER_ALIASES[Q_COL_TOPIC] = ['Lesson', 'Subtopic', 'Skill'];
HEADER_ALIASES[Q_COL_DIFFICULTY] = ['Level', 'Difficulty Level'];
HEADER_ALIASES[Q_COL_TEXT] = ['Question', 'QuestionText', 'Stem', 'Prompt'];
HEADER_ALIASES[Q_COL_OPTION_A] = ['Option A', 'Choice A', 'A'];
HEADER_ALIASES[Q_COL_OPTION_B] = ['Option B', 'Choice B', 'B'];
HEADER_ALIASES[Q_COL_OPTION_C] = ['Option C', 'Choice C', 'C'];
HEADER_ALIASES[Q_COL_OPTION_D] = ['Option D', 'Choice D', 'D'];
HEADER_ALIASES[Q_COL_IMAGE_URL] = ['Image URL', 'Image', 'FigureURL', 'Figure URL'];
HEADER_ALIASES[Q_COL_ANSWER] = ['Correct Answer', 'CorrectAnswer', 'Key', 'Answer Key'];
HEADER_ALIASES[Q_COL_EXPLANATION] = ['Rationale', 'Solution', 'Feedback'];

const USER_HEADERS = [
  'Timestamp',
  'Started At',
  'Submitted At',
  'Duration Seconds',
  'Name',
  'ID Number',
  'Total Score',
  'Total Questions',
  'Percentage',
  'Accuracy Band',
  'Category',
  'Subject',
  'Topic',
  'Difficulty',
  'Time Per Question',
  'Feedback Mode',
  'Quiz Identifier',
  'Mode',
];

const RESPONSE_HEADERS = [
  'Timestamp',
  'Name',
  'ID Number',
  'QuizIdentifier',
  'QuestionNumber',
  'QuestionID',
  'Category',
  'Subject',
  'Topic',
  'Difficulty',
  'UserAnswer',
  'CorrectAnswer',
  'IsCorrect',
  'ScoreAwarded',
  'TimedOut',
  'MarkedForReview',
  'TimeSpentSeconds',
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('QuizPage')
    .setTitle(getConfiguredAppTitle_())
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ensureSetup() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, QUESTIONS_SHEET_NAME, QUESTION_HEADERS);
  ensureSheet_(ss, USERS_SHEET_NAME, USER_HEADERS);
  ensureSheet_(ss, RESPONSES_SHEET_NAME, RESPONSE_HEADERS);
  return {
    success: true,
    message: 'Quiz app setup is ready. Existing rows were preserved. Missing sheets/headers were added only when needed.',
    appConfig: getAppConfig(),
    diagnostics: getQuestionBankStats(),
  };
}

function setSpreadsheetId(spreadsheetId) {
  const cleanId = String(spreadsheetId || '').trim();
  if (!cleanId) throw new Error('Spreadsheet ID is required.');
  PropertiesService.getScriptProperties().setProperty(SCRIPT_PROP_SPREADSHEET_ID, cleanId);
  return ensureSetup();
}

function setAppConfig(config) {
  const safeConfig = config || {};
  const props = PropertiesService.getScriptProperties();
  const title = String(safeConfig.title || '').trim();
  const subject = String(safeConfig.subject || '').trim();
  if (title) props.setProperty(SCRIPT_PROP_APP_TITLE, title);
  if (subject) props.setProperty(SCRIPT_PROP_APP_SUBJECT, subject);
  return getAppConfig();
}

function getAppConfig() {
  return {
    success: true,
    title: getConfiguredAppTitle_(),
    subject: getConfiguredSubject_(),
    schemaVersion: '5.1.0',
    maxQuestionsPerRun: MAX_QUESTIONS_PER_RUN,
    sheets: {
      questions: QUESTIONS_SHEET_NAME,
      users: USERS_SHEET_NAME,
      responses: RESPONSES_SHEET_NAME,
    },
  };
}

function getPromptTemplates() {
  return {
    success: true,
    templates: {
      fullQuestionBankGenerator: [
        'You are an expert assessment designer for [SUBJECT].',
        'Create [NUMBER] high-quality multiple-choice questions for [LEARNER LEVEL] learners.',
        'Use this exact semicolon-separated schema:',
        'Question ID;Category;Subject;Topic;Difficulty;Question Text;OptionA;OptionB;OptionC;OptionD;ImageURL;Answer;Explanation',
        'Rules: one correct answer, plausible distractors, difficulty as Easy/Normal/Difficult, answer as A/B/C/D, and an instructional explanation.',
        'Topics: [TOPICS]',
        'Special constraints: [CONSTRAINTS]',
      ].join('\n'),
      qualityReviewer: [
        'Review this question bank for answer-key correctness, duplicate items, ambiguous wording, weak distractors, and topic coverage.',
        'Return critical fixes, suggested improvements, rows with likely wrong answers, duplicates, and corrected CSV-ready rows only where needed.',
        '[PASTE QUESTIONS]',
      ].join('\n'),
      difficultyCalibrator: [
        'Calibrate the difficulty labels in this question bank.',
        'Easy = recall or one-step application. Normal = multi-step or standard exam application. Difficult = synthesis, traps, or high cognitive load.',
        'Return only rows whose difficulty should change and explain why.',
        '[PASTE QUESTIONS]',
      ].join('\n'),
      explanationImprover: [
        'Improve the explanations so learners understand why the correct answer is correct and why tempting wrong options are wrong.',
        'Keep the same schema and answer key unless the key is clearly incorrect.',
        '[PASTE QUESTIONS]',
      ].join('\n'),
    },
  };
}

function getQuestionBankStats() {
  try {
    const bank = getQuestionBank_();
    return {
      success: true,
      totalQuestions: bank.totalRows,
      validQuestions: bank.questions.length,
      invalidQuestions: bank.invalidQuestions,
      duplicateQuestionIds: bank.duplicateQuestionIds,
      invalidQuestionSamples: bank.invalidQuestionSamples,
      byCategory: countBy_(bank.questions, 'category'),
      bySubject: countBy_(bank.questions, 'subject'),
      byTopic: countBy_(bank.questions, 'topic'),
      byDifficulty: countBy_(bank.questions, 'difficulty'),
    };
  } catch (error) {
    Logger.log('Error in getQuestionBankStats: ' + error.toString() + '\nStack: ' + error.stack);
    return { success: false, error: 'Server error while reading question-bank stats: ' + error.message };
  }
}

function getInitialQuizData(filters) {
  try {
    const safeFilters = normalizeFilters_(filters || {});
    const bank = getQuestionBank_();
    const filterOptions = buildFilterOptions_(bank.questions);

    if (!bank.questions.length) {
      return {
        success: true,
        questions: [],
        filterOptions: filterOptions,
        totalAvailable: 0,
        invalidQuestions: bank.invalidQuestions,
        duplicateQuestionIds: bank.duplicateQuestionIds,
        invalidQuestionSamples: bank.invalidQuestionSamples,
        warning: 'No valid questions found yet. Add question rows below the header row.',
      };
    }

    const matched = filterQuestions_(bank.questions, safeFilters);

    if (safeFilters.setupOnly) {
      return {
        success: true,
        questions: [],
        filterOptions: filterOptions,
        totalAvailable: bank.questions.length,
        invalidQuestions: bank.invalidQuestions,
        duplicateQuestionIds: bank.duplicateQuestionIds,
        invalidQuestionSamples: bank.invalidQuestionSamples,
        matchedBeforeSlice: matched.length,
        stats: {
          byCategory: countBy_(matched, 'category'),
          bySubject: countBy_(matched, 'subject'),
          byTopic: countBy_(matched, 'topic'),
          byDifficulty: countBy_(matched, 'difficulty'),
        },
      };
    }

    if (!matched.length) {
      return {
        success: false,
        error: 'No questions match the selected filters. Adjust the Category, Subject, Topic, or Difficulty.',
        filterOptions: filterOptions,
        totalAvailable: bank.questions.length,
        invalidQuestions: bank.invalidQuestions,
      };
    }

    let questions = matched.slice();
    if (safeFilters.randomize) shuffle_(questions);
    questions = questions.slice(0, safeFilters.count);

    return {
      success: true,
      questions: questions,
      filterOptions: filterOptions,
      totalAvailable: bank.questions.length,
      invalidQuestions: bank.invalidQuestions,
      duplicateQuestionIds: bank.duplicateQuestionIds,
      matchedQuestions: questions.length,
      matchedBeforeSlice: matched.length,
    };
  } catch (error) {
    Logger.log('Error in getInitialQuizData: ' + error.toString() + '\nStack: ' + error.stack);
    return { success: false, error: 'Server error while fetching quiz data: ' + error.message };
  }
}

function getQuestionPreview(filters) {
  try {
    const safeFilters = normalizeFilters_(filters || {});
    const bank = getQuestionBank_();
    const matched = filterQuestions_(bank.questions, safeFilters);
    return {
      success: true,
      totalAvailable: bank.questions.length,
      invalidQuestions: bank.invalidQuestions,
      duplicateQuestionIds: bank.duplicateQuestionIds,
      matchedQuestions: matched.length,
      byCategory: countBy_(matched, 'category'),
      bySubject: countBy_(matched, 'subject'),
      byTopic: countBy_(matched, 'topic'),
      byDifficulty: countBy_(matched, 'difficulty'),
    };
  } catch (error) {
    Logger.log('Error in getQuestionPreview: ' + error.toString() + '\nStack: ' + error.stack);
    return { success: false, error: 'Server error while previewing filters: ' + error.message };
  }
}

function recordFullQuizResults(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const payload = normalizeResultPayload_(data || {});
    const ss = getSpreadsheet_();
    const timestamp = new Date();
    const quizIdentifier = [payload.mode.category, payload.mode.subject, payload.mode.topic, payload.mode.difficulty]
      .filter(function(part) { return part && part !== ALL_VALUE; })
      .join(' | ') || 'All Questions';

    const usersSheet = ensureSheet_(ss, USERS_SHEET_NAME, USER_HEADERS);
    const responsesSheet = ensureSheet_(ss, RESPONSES_SHEET_NAME, RESPONSE_HEADERS);
    const percentage = payload.totalQuestions > 0 ? round2_(payload.score / payload.totalQuestions * 100) : 0;

    appendObjectRows_(usersSheet, [{
      'Timestamp': timestamp,
      'Started At': payload.startedAt || '',
      'Submitted At': payload.submittedAt || timestamp,
      'Duration Seconds': payload.durationSeconds,
      'Name': payload.user.name,
      'ID Number': payload.user.idNumber,
      'Total Score': payload.score,
      'Total Questions': payload.totalQuestions,
      'Percentage': percentage,
      'Accuracy Band': getAccuracyBand_(percentage),
      'Category': payload.mode.category,
      'Subject': payload.mode.subject,
      'Topic': payload.mode.topic,
      'Difficulty': payload.mode.difficulty,
      'Time Per Question': payload.mode.timePerQuestion,
      'Feedback Mode': payload.mode.feedbackMode,
      'Quiz Identifier': quizIdentifier,
      'Mode': JSON.stringify(payload.mode),
    }]);

    if (payload.responses.length) {
      appendObjectRows_(responsesSheet, payload.responses.map(function(response, index) {
        return {
          'Timestamp': timestamp,
          'Name': payload.user.name,
          'ID Number': payload.user.idNumber,
          'QuizIdentifier': quizIdentifier,
          'QuestionNumber': response.questionNumber || index + 1,
          'QuestionID': response.questionId,
          'Category': response.category,
          'Subject': response.subject,
          'Topic': response.topic,
          'Difficulty': response.difficulty,
          'UserAnswer': response.userAnswer,
          'CorrectAnswer': response.correctAnswer,
          'IsCorrect': response.isCorrect,
          'ScoreAwarded': response.isCorrect ? 1 : 0,
          'TimedOut': response.timedOut,
          'MarkedForReview': response.markedForReview,
          'TimeSpentSeconds': response.timeSpentSeconds,
        };
      }));
    }

    return { success: true, message: 'Quiz results recorded successfully.', savedResponses: payload.responses.length };
  } catch (error) {
    Logger.log('Error in recordFullQuizResults: ' + error.toString() + '\nStack: ' + error.stack);
    return { success: false, error: 'Server error while recording quiz results: ' + error.message };
  } finally {
    try { lock.releaseLock(); } catch (releaseError) {}
  }
}

function getQuizAnalytics() {
  try {
    const ss = getSpreadsheet_();
    ensureSetup();
    const bank = getQuestionBank_();
    const questionMap = buildQuestionMap_(bank.questions);
    const attempts = normalizeAttempts_(readSheetObjects_(ss, USERS_SHEET_NAME));
    const answers = normalizeAnswers_(readSheetObjects_(ss, RESPONSES_SHEET_NAME), questionMap);
    const breakdowns = {
      byCategory: summarizeAnswerGroups_(answers, 'category'),
      bySubject: summarizeAnswerGroups_(answers, 'subject'),
      byTopic: summarizeAnswerGroups_(answers, 'topic'),
      byDifficulty: summarizeAnswerGroups_(answers, 'difficulty'),
    };
    const weakAreas = summarizeWeakAreas_(breakdowns);
    const missedQuestions = summarizeMissedQuestions_(answers);

    return {
      success: true,
      generatedAt: formatDateForClient_(new Date()),
      appConfig: getAppConfig(),
      kpis: buildKpis_(attempts, answers, bank),
      breakdowns: breakdowns,
      weakAreas: weakAreas.slice(0, 15),
      recommendations: buildRecommendations_(bank, weakAreas, missedQuestions),
      topMissedQuestions: missedQuestions.slice(0, 15),
      scoreDistribution: summarizeScoreDistribution_(attempts),
      dailyTrend: summarizeDailyTrend_(attempts).slice(-30),
      optionErrorPatterns: summarizeOptionErrorPatterns_(answers).slice(0, 15),
      studentPerformance: summarizeStudents_(attempts).slice(0, 25),
      recentAttempts: attempts.slice().sort(function(a, b) { return b.timestampMs - a.timestampMs; }).slice(0, 25),
      questionBank: {
        totalRows: bank.totalRows,
        validQuestions: bank.questions.length,
        invalidQuestions: bank.invalidQuestions,
        duplicateQuestionIds: bank.duplicateQuestionIds,
        invalidQuestionSamples: bank.invalidQuestionSamples,
        byCategory: countBy_(bank.questions, 'category'),
        bySubject: countBy_(bank.questions, 'subject'),
        byTopic: countBy_(bank.questions, 'topic'),
        byDifficulty: countBy_(bank.questions, 'difficulty'),
      },
    };
  } catch (error) {
    Logger.log('Error in getQuizAnalytics: ' + error.toString() + '\nStack: ' + error.stack);
    return { success: false, error: 'Server error while building analytics: ' + error.message };
  }
}

function getSpreadsheet_() {
  const configuredId = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_SPREADSHEET_ID) || DEFAULT_SPREADSHEET_ID;
  const spreadsheetId = String(configuredId || '').trim();
  if (!spreadsheetId || spreadsheetId === 'YOUR_SPREADSHEET_ID_HERE') {
    throw new Error('Set DEFAULT_SPREADSHEET_ID in Code.gs or run setSpreadsheetId("YOUR_SHEET_ID").');
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

function getConfiguredAppTitle_() {
  return String(PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_APP_TITLE) || APP_TITLE).trim() || APP_TITLE;
}

function getConfiguredSubject_() {
  return String(PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_APP_SUBJECT) || 'General Subject').trim() || 'General Subject';
}

function ensureSheet_(ss, sheetName, requiredHeaders) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
  } else {
    const currentHeaders = getSheetHeaders_(sheet);
    const missingHeaders = requiredHeaders.filter(function(header) { return !hasHeaderEquivalent_(currentHeaders, header); });
    if (missingHeaders.length) sheet.getRange(1, currentHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
  }

  sheet.setFrozenRows(1);
  return sheet;
}

function getSheetHeaders_(sheet) {
  if (sheet.getLastColumn() === 0) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
}

function appendObjectRows_(sheet, objects) {
  if (!objects || !objects.length) return;
  const headers = getSheetHeaders_(sheet);
  const rows = objects.map(function(object) {
    return headers.map(function(header) { return Object.prototype.hasOwnProperty.call(object, header) ? object[header] : ''; });
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function readSheetObjects_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1 || sheet.getLastColumn() === 0) return [];
  const headers = getSheetHeaders_(sheet);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues().map(function(row) {
    const item = {};
    headers.forEach(function(header, index) { item[header] = row[index]; });
    return item;
  });
}

function getQuestionBank_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(QUESTIONS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) {
    ensureSheet_(ss, QUESTIONS_SHEET_NAME, QUESTION_HEADERS);
    return buildQuestionBankResult_([], 0, []);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return buildQuestionBankResult_([], 0, []);

  const headers = values[0];
  const colMap = getColumnMap_(headers);
  const missing = getMissingRequiredQuestionColumns_(colMap);
  if (missing.length) throw new Error('Questions sheet is missing required columns: ' + missing.join(', '));

  const rows = values.slice(1);
  const validQuestions = [];
  const invalidSamples = [];

  rows.forEach(function(row, index) {
    const result = rowToQuestionResult_(row, colMap, index + 2);
    if (result.question) validQuestions.push(result.question);
    else if (invalidSamples.length < MAX_DIAGNOSTIC_ROWS) invalidSamples.push(result.error);
  });

  return buildQuestionBankResult_(validQuestions, rows.length, invalidSamples);
}

function buildQuestionBankResult_(questions, totalRows, invalidSamples) {
  const duplicates = findDuplicateValues_(questions.map(function(question) { return question.questionId; }));
  return {
    questions: questions,
    totalRows: totalRows,
    invalidQuestions: Math.max(totalRows - questions.length, 0),
    duplicateQuestionIds: duplicates,
    invalidQuestionSamples: invalidSamples || [],
  };
}

function getColumnMap_(headers) {
  const rawMap = {};
  const map = {};
  headers.forEach(function(header, index) {
    const cleanHeader = String(header || '').trim();
    if (cleanHeader && rawMap[cleanHeader] === undefined) rawMap[cleanHeader] = index;
  });

  Object.keys(rawMap).forEach(function(header) {
    map[header] = rawMap[header];
  });

  QUESTION_HEADERS.forEach(function(canonicalHeader) {
    const candidates = [canonicalHeader].concat(HEADER_ALIASES[canonicalHeader] || []);
    for (let i = 0; i < candidates.length; i++) {
      if (rawMap[candidates[i]] !== undefined) {
        map[canonicalHeader] = rawMap[candidates[i]];
        break;
      }
    }
  });

  return map;
}

function hasHeaderEquivalent_(headers, canonicalHeader) {
  const candidates = [canonicalHeader].concat(HEADER_ALIASES[canonicalHeader] || []);
  return headers.some(function(header) { return candidates.indexOf(String(header || '').trim()) !== -1; });
}

function getMissingRequiredQuestionColumns_(colMap) {
  const requiredColumns = [Q_COL_ID, Q_COL_CATEGORY, Q_COL_SUBJECT, Q_COL_TOPIC, Q_COL_TEXT, Q_COL_OPTION_A, Q_COL_OPTION_B, Q_COL_OPTION_C, Q_COL_OPTION_D, Q_COL_ANSWER, Q_COL_EXPLANATION];
  return requiredColumns.filter(function(column) { return colMap[column] === undefined; });
}

function rowToQuestion_(row, colMap) {
  return rowToQuestionResult_(row, colMap, 0).question;
}

function rowToQuestionResult_(row, colMap, rowNumber) {
  const answer = getCell_(row, colMap, Q_COL_ANSWER).toUpperCase();
  const question = {
    questionId: getCell_(row, colMap, Q_COL_ID),
    category: getCell_(row, colMap, Q_COL_CATEGORY) || 'General',
    subject: getCell_(row, colMap, Q_COL_SUBJECT) || getConfiguredSubject_(),
    topic: getCell_(row, colMap, Q_COL_TOPIC) || 'General',
    difficulty: normalizeDifficulty_(getCell_(row, colMap, Q_COL_DIFFICULTY) || 'Unspecified'),
    questionText: getCell_(row, colMap, Q_COL_TEXT),
    options: {
      A: getCell_(row, colMap, Q_COL_OPTION_A),
      B: getCell_(row, colMap, Q_COL_OPTION_B),
      C: getCell_(row, colMap, Q_COL_OPTION_C),
      D: getCell_(row, colMap, Q_COL_OPTION_D),
    },
    correctAnswer: answer,
    explanation: getCell_(row, colMap, Q_COL_EXPLANATION) || 'No explanation provided.',
    imageUrl: getCell_(row, colMap, Q_COL_IMAGE_URL),
  };

  const prefix = rowNumber ? 'Row ' + rowNumber + ': ' : '';
  if (!question.questionId) return { question: null, error: prefix + 'missing Question ID.' };
  if (!question.questionText) return { question: null, error: prefix + 'missing Question Text.' };
  if (!question.options.A || !question.options.B || !question.options.C || !question.options.D) return { question: null, error: prefix + 'all four options are required.' };
  if (['A', 'B', 'C', 'D'].indexOf(question.correctAnswer) === -1) return { question: null, error: prefix + 'Answer must be A, B, C, or D.' };

  return { question: question, error: '' };
}

function getCell_(row, colMap, columnName) {
  const index = colMap[columnName];
  if (index === undefined) return '';
  return String(row[index] === null || row[index] === undefined ? '' : row[index]).trim();
}

function buildFilterOptions_(questions) {
  return {
    categories: uniqueSorted_(questions.map(function(question) { return question.category; })),
    subjects: uniqueSorted_(questions.map(function(question) { return question.subject; })),
    topics: uniqueSorted_(questions.map(function(question) { return question.topic; })),
    difficulties: uniqueSorted_(questions.map(function(question) { return question.difficulty; })),
  };
}

function uniqueSorted_(values) {
  const seen = {};
  values.forEach(function(value) {
    const cleanValue = String(value || '').trim();
    if (cleanValue) seen[cleanValue] = true;
  });
  return Object.keys(seen).sort(function(a, b) { return a.localeCompare(b); });
}

function countBy_(items, key) {
  const counts = {};
  (items || []).forEach(function(item) {
    const value = String(item[key] || 'Unspecified').trim() || 'Unspecified';
    counts[value] = (counts[value] || 0) + 1;
  });
  return counts;
}

function filterQuestions_(questions, filters) {
  return questions.filter(function(question) {
    return matchesFilter_(question.category, filters.category)
      && matchesFilter_(question.subject, filters.subject)
      && matchesFilter_(question.topic, filters.topic)
      && matchesFilter_(question.difficulty, filters.difficulty);
  });
}

function normalizeFilters_(filters) {
  const rawCount = Number(filters.count);
  const setupOnly = Boolean(filters.setupOnly) || rawCount <= 0;
  return {
    setupOnly: setupOnly,
    category: normalizeFilterValue_(filters.category),
    subject: normalizeFilterValue_(filters.subject),
    topic: normalizeFilterValue_(filters.topic),
    difficulty: normalizeFilterValue_(filters.difficulty),
    count: setupOnly ? 0 : clamp_(rawCount || 10, 1, MAX_QUESTIONS_PER_RUN),
    randomize: filters.randomize !== false,
  };
}

function normalizeFilterValue_(value) {
  const cleanValue = String(value || ALL_VALUE).trim();
  return cleanValue || ALL_VALUE;
}

function normalizeDifficulty_(value) {
  const cleanValue = String(value || '').trim();
  const lower = cleanValue.toLowerCase();
  if (!cleanValue) return 'Unspecified';
  if (lower === 'medium' || lower === 'moderate' || lower === 'average') return 'Normal';
  if (lower === 'hard' || lower === 'advanced') return 'Difficult';
  if (lower === 'simple' || lower === 'basic') return 'Easy';
  return cleanValue;
}

function matchesFilter_(actual, expected) {
  return !expected || expected === ALL_VALUE || String(actual || '').trim() === expected;
}

function normalizeResultPayload_(data) {
  const userDetails = data.userDetails || {};
  const quizSummary = data.quizSummary || {};
  const mode = quizSummary.mode || {};
  const responses = Array.isArray(data.responses) ? data.responses : [];
  const score = clamp_(Number(quizSummary.score) || 0, 0, Number(quizSummary.totalQuestions) || responses.length || 0);
  const totalQuestions = Math.max(Number(quizSummary.totalQuestions) || responses.length || 0, responses.length);

  const normalized = {
    user: {
      name: String(userDetails.name || '').trim(),
      idNumber: String(userDetails.idNumber || '').trim(),
    },
    score: score,
    totalQuestions: totalQuestions,
    startedAt: normalizeDateString_(data.startedAt),
    submittedAt: normalizeDateString_(data.submittedAt) || new Date(),
    durationSeconds: clamp_(Number(data.durationSeconds) || 0, 0, 86400),
    mode: {
      numQuestions: Number(mode.numQuestions) || totalQuestions,
      timePerQuestion: String(mode.timePerQuestion || '').trim(),
      category: normalizeFilterValue_(mode.category),
      subject: normalizeFilterValue_(mode.subject),
      topic: normalizeFilterValue_(mode.topic),
      difficulty: normalizeFilterValue_(mode.difficulty),
      feedbackMode: String(mode.feedbackMode || 'Instant feedback').trim(),
    },
    responses: responses.map(normalizeResponse_).filter(Boolean),
  };

  if (!normalized.user.name) throw new Error('Name is required before saving results.');
  if (!normalized.user.idNumber) throw new Error('ID Number is required before saving results.');
  return normalized;
}

function normalizeResponse_(response) {
  if (!response || !response.questionId) return null;
  const userAnswer = String(response.userAnswer || '').trim().toUpperCase() || 'TIMEOUT';
  return {
    questionNumber: clamp_(Number(response.questionNumber) || 0, 0, MAX_QUESTIONS_PER_RUN),
    questionId: String(response.questionId || '').trim(),
    category: String(response.category || '').trim(),
    subject: String(response.subject || '').trim(),
    topic: String(response.topic || '').trim(),
    difficulty: String(response.difficulty || '').trim(),
    userAnswer: userAnswer,
    correctAnswer: String(response.correctAnswer || '').trim().toUpperCase(),
    isCorrect: Boolean(response.isCorrect),
    timedOut: Boolean(response.timedOut),
    markedForReview: Boolean(response.markedForReview),
    timeSpentSeconds: clamp_(Number(response.timeSpentSeconds) || 0, 0, 86400),
  };
}

function buildQuestionMap_(questions) {
  const map = {};
  questions.forEach(function(question) { map[question.questionId] = question; });
  return map;
}

function normalizeAttempts_(rows) {
  return (rows || []).map(function(row) {
    const scoreInfo = parseScore_(row['Total Score'], row['Total Questions']);
    const timestampDate = normalizeDateString_(row['Submitted At'] || row['Timestamp']);
    const name = String(row['Name'] || '').trim() || 'Unknown';
    const idNumber = String(row['ID Number'] || '').trim();
    const percentage = round2_(toNumber_(row['Percentage'], scoreInfo.total ? scoreInfo.score / scoreInfo.total * 100 : 0));
    return {
      timestamp: formatDateForClient_(timestampDate),
      dateKey: formatDateKey_(timestampDate),
      timestampMs: timestampDate ? timestampDate.getTime() : 0,
      student: maskName_(name) + (idNumber ? ' (' + maskId_(idNumber) + ')' : ''),
      studentKey: name + '|' + idNumber,
      score: scoreInfo.score,
      totalQuestions: scoreInfo.total,
      percentage: percentage,
      durationSeconds: toNumber_(row['Duration Seconds'], 0),
      category: String(row['Category'] || ALL_VALUE),
      subject: String(row['Subject'] || ALL_VALUE),
      topic: String(row['Topic'] || ALL_VALUE),
      difficulty: String(row['Difficulty'] || ALL_VALUE),
      feedbackMode: String(row['Feedback Mode'] || ''),
    };
  }).filter(function(item) { return item.studentKey.trim(); });
}

function normalizeAnswers_(rows, questionMap) {
  return (rows || []).map(function(row) {
    const questionId = String(row['QuestionID'] || '').trim();
    const question = questionMap[questionId] || {};
    const timestampDate = normalizeDateString_(row['Timestamp']);
    return {
      timestamp: formatDateForClient_(timestampDate),
      dateKey: formatDateKey_(timestampDate),
      timestampMs: timestampDate ? timestampDate.getTime() : 0,
      questionNumber: toNumber_(row['QuestionNumber'], 0),
      questionId: questionId,
      questionText: question.questionText || '',
      category: String(row['Category'] || question.category || 'Unspecified'),
      subject: String(row['Subject'] || question.subject || 'Unspecified'),
      topic: String(row['Topic'] || question.topic || 'Unspecified'),
      difficulty: String(row['Difficulty'] || question.difficulty || 'Unspecified'),
      userAnswer: String(row['UserAnswer'] || '').trim().toUpperCase(),
      correctAnswer: String(row['CorrectAnswer'] || question.correctAnswer || '').trim().toUpperCase(),
      isCorrect: toBoolean_(row['IsCorrect']),
      timedOut: toBoolean_(row['TimedOut']),
      markedForReview: toBoolean_(row['MarkedForReview']),
      timeSpentSeconds: toNumber_(row['TimeSpentSeconds'], 0),
    };
  }).filter(function(item) { return item.questionId; });
}

function buildKpis_(attempts, answers, bank) {
  const correctAnswers = answers.filter(function(answer) { return answer.isCorrect; }).length;
  const timedOut = answers.filter(function(answer) { return answer.timedOut; }).length;
  const marked = answers.filter(function(answer) { return answer.markedForReview; }).length;
  const percentages = attempts.map(function(attempt) { return attempt.percentage; });
  const durationValues = attempts.map(function(attempt) { return attempt.durationSeconds; }).filter(function(value) { return value > 0; });
  return {
    attempts: attempts.length,
    uniqueStudents: uniqueCount_(attempts.map(function(attempt) { return attempt.studentKey; })),
    totalAnswers: answers.length,
    avgAttemptPercentage: attempts.length ? round2_(average_(percentages)) : 0,
    medianAttemptPercentage: attempts.length ? round2_(median_(percentages)) : 0,
    answerAccuracy: answers.length ? round2_(correctAnswers / answers.length * 100) : 0,
    avgDurationSeconds: durationValues.length ? round2_(average_(durationValues)) : 0,
    avgQuestionsPerAttempt: attempts.length ? round2_(average_(attempts.map(function(attempt) { return attempt.totalQuestions; }))) : 0,
    timeoutRate: answers.length ? round2_(timedOut / answers.length * 100) : 0,
    markedForReviewRate: answers.length ? round2_(marked / answers.length * 100) : 0,
    validQuestions: bank.questions.length,
    invalidQuestionRows: bank.invalidQuestions,
    duplicateQuestionIds: bank.duplicateQuestionIds.length,
  };
}

function summarizeAnswerGroups_(answers, key) {
  const groups = {};
  answers.forEach(function(answer) {
    const label = String(answer[key] || 'Unspecified').trim() || 'Unspecified';
    if (!groups[label]) groups[label] = { label: label, answers: 0, correct: 0, incorrect: 0, timedOut: 0, marked: 0, totalTime: 0, timedAnswers: 0 };
    groups[label].answers++;
    if (answer.isCorrect) groups[label].correct++;
    else groups[label].incorrect++;
    if (answer.timedOut) groups[label].timedOut++;
    if (answer.markedForReview) groups[label].marked++;
    if (answer.timeSpentSeconds > 0) {
      groups[label].totalTime += answer.timeSpentSeconds;
      groups[label].timedAnswers++;
    }
  });

  return Object.keys(groups).map(function(label) {
    const group = groups[label];
    const accuracy = group.answers ? group.correct / group.answers * 100 : 0;
    return {
      label: label,
      answers: group.answers,
      correct: group.correct,
      incorrect: group.incorrect,
      accuracy: round2_(accuracy),
      missRate: group.answers ? round2_(group.incorrect / group.answers * 100) : 0,
      timeoutRate: group.answers ? round2_(group.timedOut / group.answers * 100) : 0,
      markedRate: group.answers ? round2_(group.marked / group.answers * 100) : 0,
      avgTimeSeconds: group.timedAnswers ? round2_(group.totalTime / group.timedAnswers) : 0,
      priorityScore: round2_((group.incorrect * 2) + group.timedOut + group.marked + Math.max(0, 70 - accuracy)),
    };
  }).sort(function(a, b) { return b.answers - a.answers || a.label.localeCompare(b.label); });
}

function summarizeWeakAreas_(breakdowns) {
  const rows = [];
  ['byCategory', 'bySubject', 'byTopic', 'byDifficulty'].forEach(function(groupKey) {
    const type = groupKey.replace('by', '');
    (breakdowns[groupKey] || []).forEach(function(row) {
      if (row.answers < 3) return;
      if (row.accuracy >= 75 && row.timeoutRate < 20 && row.markedRate < 20) return;
      rows.push({
        type: type,
        label: row.label,
        answers: row.answers,
        accuracy: row.accuracy,
        missRate: row.missRate,
        timeoutRate: row.timeoutRate,
        markedRate: row.markedRate,
        avgTimeSeconds: row.avgTimeSeconds,
        priorityScore: row.priorityScore,
      });
    });
  });
  return rows.sort(function(a, b) { return b.priorityScore - a.priorityScore || b.answers - a.answers; });
}

function summarizeMissedQuestions_(answers) {
  const groups = {};
  answers.forEach(function(answer) {
    if (!groups[answer.questionId]) {
      groups[answer.questionId] = {
        questionId: answer.questionId,
        questionText: answer.questionText,
        category: answer.category,
        subject: answer.subject,
        topic: answer.topic,
        difficulty: answer.difficulty,
        attempts: 0,
        misses: 0,
        timedOut: 0,
        marked: 0,
        totalTime: 0,
        timedAnswers: 0,
        correctAnswer: answer.correctAnswer,
        wrongAnswers: {},
      };
    }
    const group = groups[answer.questionId];
    group.attempts++;
    if (!answer.isCorrect) {
      group.misses++;
      const wrong = answer.userAnswer || (answer.timedOut ? 'TIMEOUT' : 'BLANK');
      group.wrongAnswers[wrong] = (group.wrongAnswers[wrong] || 0) + 1;
    }
    if (answer.timedOut) group.timedOut++;
    if (answer.markedForReview) group.marked++;
    if (answer.timeSpentSeconds > 0) {
      group.totalTime += answer.timeSpentSeconds;
      group.timedAnswers++;
    }
  });

  return Object.keys(groups).map(function(questionId) {
    const group = groups[questionId];
    group.missRate = group.attempts ? round2_(group.misses / group.attempts * 100) : 0;
    group.timeoutRate = group.attempts ? round2_(group.timedOut / group.attempts * 100) : 0;
    group.markedRate = group.attempts ? round2_(group.marked / group.attempts * 100) : 0;
    group.avgTimeSeconds = group.timedAnswers ? round2_(group.totalTime / group.timedAnswers) : 0;
    group.commonWrongAnswer = topKey_(group.wrongAnswers);
    return group;
  }).filter(function(group) { return group.misses > 0; }).sort(function(a, b) {
    return b.missRate - a.missRate || b.misses - a.misses || b.markedRate - a.markedRate;
  });
}

function summarizeScoreDistribution_(attempts) {
  const buckets = [
    { label: '0-49%', min: 0, max: 49.999, count: 0 },
    { label: '50-69%', min: 50, max: 69.999, count: 0 },
    { label: '70-84%', min: 70, max: 84.999, count: 0 },
    { label: '85-100%', min: 85, max: 100, count: 0 },
  ];
  attempts.forEach(function(attempt) {
    const bucket = buckets.filter(function(item) { return attempt.percentage >= item.min && attempt.percentage <= item.max; })[0] || buckets[0];
    bucket.count++;
  });
  return buckets.map(function(bucket) {
    return { label: bucket.label, count: bucket.count, percentage: attempts.length ? round2_(bucket.count / attempts.length * 100) : 0 };
  });
}

function summarizeDailyTrend_(attempts) {
  const groups = {};
  attempts.forEach(function(attempt) {
    const key = attempt.dateKey || 'Unknown Date';
    if (!groups[key]) groups[key] = { date: key, attempts: 0, totalPercentage: 0, totalDuration: 0, durationCount: 0 };
    groups[key].attempts++;
    groups[key].totalPercentage += attempt.percentage;
    if (attempt.durationSeconds > 0) {
      groups[key].totalDuration += attempt.durationSeconds;
      groups[key].durationCount++;
    }
  });
  return Object.keys(groups).sort().map(function(key) {
    const group = groups[key];
    return {
      date: group.date,
      attempts: group.attempts,
      avgPercentage: round2_(group.totalPercentage / group.attempts),
      avgDurationSeconds: group.durationCount ? round2_(group.totalDuration / group.durationCount) : 0,
    };
  });
}

function summarizeOptionErrorPatterns_(answers) {
  const groups = {};
  answers.forEach(function(answer) {
    if (answer.isCorrect) return;
    const selected = answer.userAnswer || (answer.timedOut ? 'TIMEOUT' : 'BLANK');
    const key = answer.questionId + '|' + selected;
    if (!groups[key]) groups[key] = { questionId: answer.questionId, selectedAnswer: selected, correctAnswer: answer.correctAnswer, topic: answer.topic, difficulty: answer.difficulty, count: 0 };
    groups[key].count++;
  });
  return Object.keys(groups).map(function(key) { return groups[key]; }).sort(function(a, b) { return b.count - a.count; });
}

function summarizeStudents_(attempts) {
  const groups = {};
  attempts.sort(function(a, b) { return a.timestampMs - b.timestampMs; }).forEach(function(attempt) {
    const key = attempt.studentKey;
    if (!groups[key]) groups[key] = { student: attempt.student, attempts: 0, best: 0, total: 0, first: null, latest: null };
    groups[key].attempts++;
    groups[key].best = Math.max(groups[key].best, attempt.percentage);
    groups[key].total += attempt.percentage;
    if (!groups[key].first || attempt.timestampMs < groups[key].first.timestampMs) groups[key].first = attempt;
    if (!groups[key].latest || attempt.timestampMs >= groups[key].latest.timestampMs) groups[key].latest = attempt;
  });

  return Object.keys(groups).map(function(key) {
    const group = groups[key];
    const first = group.first || {};
    const latest = group.latest || {};
    return {
      student: group.student,
      attempts: group.attempts,
      bestPercentage: round2_(group.best),
      averagePercentage: round2_(group.total / group.attempts),
      latestPercentage: round2_(latest.percentage || 0),
      improvement: round2_((latest.percentage || 0) - (first.percentage || 0)),
      latestAttempt: latest.timestamp || '',
    };
  }).sort(function(a, b) { return b.averagePercentage - a.averagePercentage; });
}

function buildRecommendations_(bank, weakAreas, missedQuestions) {
  const recommendations = [];
  if (bank.invalidQuestions) recommendations.push('Fix invalid question rows before high-stakes use. Open setup diagnostics to inspect missing fields or invalid answer keys.');
  if (bank.duplicateQuestionIds.length) recommendations.push('Rename duplicate Question IDs so analytics can map responses to the correct item.');
  if (weakAreas.length) recommendations.push('Prioritize review for: ' + weakAreas.slice(0, 3).map(function(item) { return item.type + ' ' + item.label; }).join(', ') + '.');
  if (missedQuestions.length) recommendations.push('Review the top missed questions and improve explanations or distractors if the miss rate is unexpectedly high.');
  if (!recommendations.length) recommendations.push('Question bank and attempt analytics look healthy. Continue adding balanced questions across topics and difficulty levels.');
  return recommendations;
}

function parseScore_(scoreValue, totalValue) {
  const total = Number(totalValue) || 0;
  if (typeof scoreValue === 'string' && scoreValue.indexOf('/') !== -1) {
    const parts = scoreValue.split('/');
    return { score: Number(parts[0]) || 0, total: Number(parts[1]) || total || 0 };
  }
  return { score: Number(scoreValue) || 0, total: total };
}

function toBoolean_(value) {
  if (typeof value === 'boolean') return value;
  const text = String(value || '').trim().toLowerCase();
  return text === 'true' || text === 'yes' || text === '1';
}

function toNumber_(value, fallback) {
  const numeric = Number(value);
  return isNaN(numeric) ? fallback : numeric;
}

function average_(values) {
  const clean = values.filter(function(value) { return typeof value === 'number' && !isNaN(value); });
  return clean.length ? clean.reduce(function(sum, value) { return sum + value; }, 0) / clean.length : 0;
}

function median_(values) {
  const clean = values.filter(function(value) { return typeof value === 'number' && !isNaN(value); }).sort(function(a, b) { return a - b; });
  if (!clean.length) return 0;
  const mid = Math.floor(clean.length / 2);
  return clean.length % 2 ? clean[mid] : (clean[mid - 1] + clean[mid]) / 2;
}

function uniqueCount_(values) {
  const seen = {};
  values.forEach(function(value) { if (value) seen[value] = true; });
  return Object.keys(seen).length;
}

function findDuplicateValues_(values) {
  const counts = {};
  values.forEach(function(value) {
    const clean = String(value || '').trim();
    if (clean) counts[clean] = (counts[clean] || 0) + 1;
  });
  return Object.keys(counts).filter(function(value) { return counts[value] > 1; }).sort();
}

function normalizeDateString_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') return isNaN(value.getTime()) ? '' : value;
  const date = new Date(value);
  return isNaN(date.getTime()) ? '' : date;
}

function formatDateForClient_(value) {
  const date = normalizeDateString_(value);
  return date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '';
}

function formatDateKey_(value) {
  const date = normalizeDateString_(value);
  return date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'Unknown Date';
}

function maskName_(name) {
  const parts = String(name || '').trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return 'Unknown';
  if (parts.length === 1) return parts[0].charAt(0) + '***';
  return parts[0].charAt(0) + '*** ' + parts[parts.length - 1].charAt(0) + '***';
}

function maskId_(idNumber) {
  const value = String(idNumber || '').trim();
  if (value.length <= 4) return '****';
  return '***' + value.slice(-4);
}

function topKey_(counts) {
  const keys = Object.keys(counts || {});
  if (!keys.length) return '';
  keys.sort(function(a, b) { return counts[b] - counts[a] || a.localeCompare(b); });
  return keys[0] + ' (' + counts[keys[0]] + ')';
}

function getAccuracyBand_(percentage) {
  if (percentage >= 85) return 'Excellent';
  if (percentage >= 70) return 'Good';
  if (percentage >= 50) return 'Needs Review';
  return 'Needs Practice';
}

function round2_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

function clamp_(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function shuffle_(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}
