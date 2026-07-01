/**
 * Civil Engineering Quiz App - Hardened Backend
 *
 * Copy this single file into Apps Script as Code.gs. The app uses one Google
 * Sheet as its database and serves QuizPage.html as the student quiz portal.
 *
 * @version 4.0.0
 * @license MIT
 */

const APP_TITLE = 'Civil Engineering Quiz App';
const DEFAULT_SPREADSHEET_ID = '1qQw7B6sRrTkbGPViBmYqqaJEqwQr-7P0jeNuPirrgpY';
const SCRIPT_PROP_SPREADSHEET_ID = 'SPREADSHEET_ID';

const QUESTIONS_SHEET_NAME = 'Questions';
const RESPONSES_SHEET_NAME = 'Responses';
const USERS_SHEET_NAME = 'Users';

const ALL_VALUE = 'All';
const MAX_QUESTIONS_PER_RUN = 100;

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

const USER_HEADERS = [
  'Timestamp',
  'Name',
  'ID Number',
  'Total Score',
  'Total Questions',
  'Percentage',
  'Category',
  'Subject',
  'Topic',
  'Difficulty',
  'Time Per Question',
  'Quiz Identifier',
  'Mode',
];

const RESPONSE_HEADERS = [
  'Timestamp',
  'Name',
  'ID Number',
  'QuizIdentifier',
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
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('QuizPage')
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Optional setup helper. Run this once after copying the files into Apps Script.
 * It creates the required sheets and headers without deleting existing data.
 */
function ensureSetup() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, QUESTIONS_SHEET_NAME, QUESTION_HEADERS);
  ensureSheet_(ss, USERS_SHEET_NAME, USER_HEADERS);
  ensureSheet_(ss, RESPONSES_SHEET_NAME, RESPONSE_HEADERS);

  return {
    success: true,
    message: 'Quiz app setup is ready. Add questions to the Questions sheet, then deploy the web app.',
  };
}

/**
 * Optional helper for non-technical deployments. Run once with your Sheet ID if
 * you do not want to edit DEFAULT_SPREADSHEET_ID directly.
 */
function setSpreadsheetId(spreadsheetId) {
  const cleanId = String(spreadsheetId || '').trim();
  if (!cleanId) {
    throw new Error('Spreadsheet ID is required.');
  }
  PropertiesService.getScriptProperties().setProperty(SCRIPT_PROP_SPREADSHEET_ID, cleanId);
  return ensureSetup();
}

/**
 * Gets filter options and, when requested, a randomized question set.
 */
function getInitialQuizData(filters) {
  try {
    const safeFilters = normalizeFilters_(filters || {});
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(QUESTIONS_SHEET_NAME);

    if (!sheet || sheet.getLastRow() < 1) {
      ensureSheet_(ss, QUESTIONS_SHEET_NAME, QUESTION_HEADERS);
      return {
        success: true,
        questions: [],
        filterOptions: buildEmptyFilterOptions_(),
        totalAvailable: 0,
        warning: 'Questions sheet is ready, but no questions have been added yet.',
      };
    }

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) {
      return {
        success: true,
        questions: [],
        filterOptions: buildEmptyFilterOptions_(),
        totalAvailable: 0,
        warning: 'No questions found yet. Add question rows below the header row.',
      };
    }

    const headers = values[0];
    const colMap = getColumnMap_(headers);
    const missing = getMissingRequiredQuestionColumns_(colMap);
    if (missing.length) {
      return {
        success: false,
        error: 'Questions sheet is missing required columns: ' + missing.join(', '),
      };
    }

    const questions = values
      .slice(1)
      .map(function(row) {
        return rowToQuestion_(row, colMap);
      })
      .filter(Boolean);

    const filterOptions = buildFilterOptions_(questions);

    if (safeFilters.setupOnly) {
      return {
        success: true,
        questions: [],
        filterOptions: filterOptions,
        totalAvailable: questions.length,
      };
    }

    let filtered = questions.filter(function(question) {
      return matchesFilter_(question.category, safeFilters.category)
        && matchesFilter_(question.subject, safeFilters.subject)
        && matchesFilter_(question.topic, safeFilters.topic)
        && matchesFilter_(question.difficulty, safeFilters.difficulty);
    });

    if (!filtered.length) {
      return {
        success: false,
        error: 'No questions match the selected filters. Adjust the Category, Subject, Topic, or Difficulty.',
        filterOptions: filterOptions,
        totalAvailable: questions.length,
      };
    }

    if (safeFilters.randomize) {
      shuffle_(filtered);
    }

    filtered = filtered.slice(0, safeFilters.count);

    return {
      success: true,
      questions: filtered,
      filterOptions: filterOptions,
      totalAvailable: questions.length,
      matchedQuestions: filtered.length,
    };
  } catch (error) {
    Logger.log('Error in getInitialQuizData: ' + error.toString() + '\nStack: ' + error.stack);
    return {
      success: false,
      error: 'Server error while fetching quiz data: ' + error.message,
    };
  }
}

/**
 * Records the final quiz summary and every per-question answer in batch writes.
 */
function recordFullQuizResults(data) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(20000);

    const payload = normalizeResultPayload_(data || {});
    const ss = getSpreadsheet_();
    const timestamp = new Date();
    const quizIdentifier = [
      payload.mode.category,
      payload.mode.subject,
      payload.mode.topic,
      payload.mode.difficulty,
    ].filter(function(part) {
      return part && part !== ALL_VALUE;
    }).join(' | ') || 'All Questions';

    const usersSheet = ensureSheet_(ss, USERS_SHEET_NAME, USER_HEADERS);
    const responsesSheet = ensureSheet_(ss, RESPONSES_SHEET_NAME, RESPONSE_HEADERS);
    const percentage = payload.totalQuestions > 0
      ? Math.round((payload.score / payload.totalQuestions) * 10000) / 100
      : 0;

    appendObjectRows_(usersSheet, [{
      'Timestamp': timestamp,
      'Name': payload.user.name,
      'ID Number': payload.user.idNumber,
      'Total Score': payload.score,
      'Total Questions': payload.totalQuestions,
      'Percentage': percentage,
      'Category': payload.mode.category,
      'Subject': payload.mode.subject,
      'Topic': payload.mode.topic,
      'Difficulty': payload.mode.difficulty,
      'Time Per Question': payload.mode.timePerQuestion,
      'Quiz Identifier': quizIdentifier,
      'Mode': JSON.stringify(payload.mode),
    }]);

    if (payload.responses.length) {
      const responseRows = payload.responses.map(function(response) {
        return {
          'Timestamp': timestamp,
          'Name': payload.user.name,
          'ID Number': payload.user.idNumber,
          'QuizIdentifier': quizIdentifier,
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
        };
      });
      appendObjectRows_(responsesSheet, responseRows);
    }

    return {
      success: true,
      message: 'Quiz results recorded successfully.',
      savedResponses: payload.responses.length,
    };
  } catch (error) {
    Logger.log('Error in recordFullQuizResults: ' + error.toString() + '\nStack: ' + error.stack);
    return {
      success: false,
      error: 'Server error while recording quiz results: ' + error.message,
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseError) {
      // Lock was not acquired or was already released.
    }
  }
}

function getSpreadsheet_() {
  const configuredId = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_SPREADSHEET_ID)
    || DEFAULT_SPREADSHEET_ID;
  const spreadsheetId = String(configuredId || '').trim();

  if (!spreadsheetId || spreadsheetId === 'YOUR_SPREADSHEET_ID_HERE') {
    throw new Error('Set DEFAULT_SPREADSHEET_ID in Code.gs or run setSpreadsheetId("YOUR_SHEET_ID").');
  }

  return SpreadsheetApp.openById(spreadsheetId);
}

function ensureSheet_(ss, sheetName, requiredHeaders) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
  } else {
    const currentHeaders = getSheetHeaders_(sheet);
    const missingHeaders = requiredHeaders.filter(function(header) {
      return currentHeaders.indexOf(header) === -1;
    });

    if (missingHeaders.length) {
      sheet.getRange(1, currentHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    }
  }

  sheet.setFrozenRows(1);
  return sheet;
}

function getSheetHeaders_(sheet) {
  if (sheet.getLastColumn() === 0) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(header) {
      return String(header || '').trim();
    });
}

function appendObjectRows_(sheet, objects) {
  if (!objects || !objects.length) return;

  const headers = getSheetHeaders_(sheet);
  const rows = objects.map(function(object) {
    return headers.map(function(header) {
      return Object.prototype.hasOwnProperty.call(object, header) ? object[header] : '';
    });
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function getColumnMap_(headers) {
  const map = {};
  headers.forEach(function(header, index) {
    const cleanHeader = String(header || '').trim();
    if (cleanHeader) {
      map[cleanHeader] = index;
    }
  });
  return map;
}

function getMissingRequiredQuestionColumns_(colMap) {
  const requiredColumns = [
    Q_COL_ID,
    Q_COL_CATEGORY,
    Q_COL_SUBJECT,
    Q_COL_TOPIC,
    Q_COL_TEXT,
    Q_COL_OPTION_A,
    Q_COL_OPTION_B,
    Q_COL_OPTION_C,
    Q_COL_OPTION_D,
    Q_COL_ANSWER,
    Q_COL_EXPLANATION,
  ];

  return requiredColumns.filter(function(column) {
    return colMap[column] === undefined;
  });
}

function rowToQuestion_(row, colMap) {
  const answer = getCell_(row, colMap, Q_COL_ANSWER).toUpperCase();
  const question = {
    questionId: getCell_(row, colMap, Q_COL_ID),
    category: getCell_(row, colMap, Q_COL_CATEGORY),
    subject: getCell_(row, colMap, Q_COL_SUBJECT),
    topic: getCell_(row, colMap, Q_COL_TOPIC),
    difficulty: getCell_(row, colMap, Q_COL_DIFFICULTY) || 'Unspecified',
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

  if (!question.questionId || !question.questionText) return null;
  if (!question.options.A || !question.options.B || !question.options.C || !question.options.D) return null;
  if (['A', 'B', 'C', 'D'].indexOf(question.correctAnswer) === -1) return null;

  return question;
}

function getCell_(row, colMap, columnName) {
  const index = colMap[columnName];
  if (index === undefined) return '';
  return String(row[index] === null || row[index] === undefined ? '' : row[index]).trim();
}

function buildEmptyFilterOptions_() {
  return {
    categories: [],
    subjects: [],
    topics: [],
    difficulties: [],
  };
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
    if (cleanValue) {
      seen[cleanValue] = true;
    }
  });

  return Object.keys(seen).sort(function(a, b) {
    return a.localeCompare(b);
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
    mode: {
      numQuestions: Number(mode.numQuestions) || totalQuestions,
      timePerQuestion: String(mode.timePerQuestion || '').trim(),
      category: normalizeFilterValue_(mode.category),
      subject: normalizeFilterValue_(mode.subject),
      topic: normalizeFilterValue_(mode.topic),
      difficulty: normalizeFilterValue_(mode.difficulty),
    },
    responses: responses.map(normalizeResponse_).filter(Boolean),
  };

  if (!normalized.user.name) {
    throw new Error('Name is required before saving results.');
  }
  if (!normalized.user.idNumber) {
    throw new Error('ID Number is required before saving results.');
  }

  return normalized;
}

function normalizeResponse_(response) {
  if (!response || !response.questionId) return null;

  const userAnswer = String(response.userAnswer || '').trim().toUpperCase() || 'TIMEOUT';
  const correctAnswer = String(response.correctAnswer || '').trim().toUpperCase();

  return {
    questionId: String(response.questionId || '').trim(),
    category: String(response.category || '').trim(),
    subject: String(response.subject || '').trim(),
    topic: String(response.topic || '').trim(),
    difficulty: String(response.difficulty || '').trim(),
    userAnswer: userAnswer,
    correctAnswer: correctAnswer,
    isCorrect: Boolean(response.isCorrect),
    timedOut: Boolean(response.timedOut),
  };
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
