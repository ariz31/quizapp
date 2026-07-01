/**
 * Optional analytics helpers for the Civil Engineering Quiz App.
 *
 * Copy this file into Apps Script as Analytics.gs when you want the faculty
 * analytics dashboard in QuizPage.html. The functions only read existing sheets
 * and do not modify stored quiz data.
 *
 * @version 1.0.0
 * @license MIT
 */

function getQuestionPreview(filters) {
  try {
    const bank = analyticsQuestionBank_();
    const safeFilters = analyticsNormalizeFilters_(filters || {});
    const matched = bank.questions.filter(function(question) {
      return analyticsMatches_(question.category, safeFilters.category)
        && analyticsMatches_(question.subject, safeFilters.subject)
        && analyticsMatches_(question.topic, safeFilters.topic)
        && analyticsMatches_(question.difficulty, safeFilters.difficulty);
    });

    return {
      success: true,
      totalAvailable: bank.questions.length,
      invalidQuestions: bank.invalidQuestions,
      matchedQuestions: matched.length,
      byCategory: analyticsCountBy_(matched, 'category'),
      bySubject: analyticsCountBy_(matched, 'subject'),
      byTopic: analyticsCountBy_(matched, 'topic'),
      byDifficulty: analyticsCountBy_(matched, 'difficulty'),
    };
  } catch (error) {
    Logger.log('getQuestionPreview failed: ' + error.toString() + '\n' + error.stack);
    return { success: false, error: 'Unable to build the question preview: ' + error.message };
  }
}

function getQuizAnalytics() {
  try {
    const ss = SpreadsheetApp.openById(analyticsSpreadsheetId_());
    const users = analyticsReadRows_(ss, USERS_SHEET_NAME);
    const responses = analyticsReadRows_(ss, RESPONSES_SHEET_NAME);
    const bank = analyticsQuestionBank_();
    const questionMap = analyticsQuestionMap_(bank.questions);

    const attempts = analyticsNormalizeAttempts_(users);
    const answers = analyticsNormalizeAnswers_(responses, questionMap);

    return {
      success: true,
      generatedAt: analyticsFormatDate_(new Date()),
      kpis: analyticsKpis_(attempts, answers, bank),
      breakdowns: {
        byCategory: analyticsAnswerGroups_(answers, 'category'),
        bySubject: analyticsAnswerGroups_(answers, 'subject'),
        byTopic: analyticsAnswerGroups_(answers, 'topic'),
        byDifficulty: analyticsAnswerGroups_(answers, 'difficulty'),
      },
      topMissedQuestions: analyticsMissedQuestions_(answers).slice(0, 15),
      studentPerformance: analyticsStudents_(attempts).slice(0, 25),
      recentAttempts: attempts.sort(function(a, b) { return b.timestampMs - a.timestampMs; }).slice(0, 25),
      questionBank: {
        totalRows: bank.totalRows,
        validQuestions: bank.questions.length,
        invalidQuestions: bank.invalidQuestions,
        byCategory: analyticsCountBy_(bank.questions, 'category'),
        byDifficulty: analyticsCountBy_(bank.questions, 'difficulty'),
      },
    };
  } catch (error) {
    Logger.log('getQuizAnalytics failed: ' + error.toString() + '\n' + error.stack);
    return { success: false, error: 'Unable to build analytics: ' + error.message };
  }
}

function analyticsSpreadsheetId_() {
  const configuredId = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_SPREADSHEET_ID) || DEFAULT_SPREADSHEET_ID;
  const spreadsheetId = String(configuredId || '').trim();
  if (!spreadsheetId || spreadsheetId === 'YOUR_SPREADSHEET_ID_HERE') {
    throw new Error('Set DEFAULT_SPREADSHEET_ID in Code.gs or run setSpreadsheetId("YOUR_SHEET_ID").');
  }
  return spreadsheetId;
}

function analyticsQuestionBank_() {
  const ss = SpreadsheetApp.openById(analyticsSpreadsheetId_());
  const sheet = ss.getSheetByName(QUESTIONS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { questions: [], totalRows: 0, invalidQuestions: 0 };
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(header) { return String(header || '').trim(); });
  const colMap = analyticsColumnMap_(headers);
  const rows = values.slice(1);
  const questions = rows.map(function(row) { return analyticsQuestionFromRow_(row, colMap); }).filter(Boolean);
  return { questions: questions, totalRows: rows.length, invalidQuestions: rows.length - questions.length };
}

function analyticsQuestionFromRow_(row, colMap) {
  const answer = analyticsCell_(row, colMap, Q_COL_ANSWER).toUpperCase();
  const question = {
    questionId: analyticsCell_(row, colMap, Q_COL_ID),
    category: analyticsCell_(row, colMap, Q_COL_CATEGORY),
    subject: analyticsCell_(row, colMap, Q_COL_SUBJECT),
    topic: analyticsCell_(row, colMap, Q_COL_TOPIC),
    difficulty: analyticsCell_(row, colMap, Q_COL_DIFFICULTY) || 'Unspecified',
    questionText: analyticsCell_(row, colMap, Q_COL_TEXT),
    correctAnswer: answer,
  };

  if (!question.questionId || !question.questionText) return null;
  if (['A', 'B', 'C', 'D'].indexOf(question.correctAnswer) === -1) return null;
  return question;
}

function analyticsReadRows_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1 || sheet.getLastColumn() === 0) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues().map(function(row) {
    const item = {};
    headers.forEach(function(header, index) { item[header] = row[index]; });
    return item;
  });
}

function analyticsNormalizeAttempts_(rows) {
  return rows.map(function(row) {
    const score = analyticsScore_(row['Total Score'], row['Total Questions']);
    const timestamp = analyticsDate_(row['Submitted At'] || row['Timestamp']);
    const rawName = String(row['Name'] || '').trim() || 'Unknown';
    const rawId = String(row['ID Number'] || '').trim();
    return {
      timestamp: analyticsFormatDate_(timestamp),
      timestampMs: timestamp ? timestamp.getTime() : 0,
      student: analyticsMaskName_(rawName) + (rawId ? ' (' + analyticsMaskId_(rawId) + ')' : ''),
      studentKey: rawName + '|' + rawId,
      score: score.score,
      totalQuestions: score.total,
      percentage: analyticsRound_(analyticsNumber_(row['Percentage'], score.total ? score.score / score.total * 100 : 0)),
      durationSeconds: analyticsNumber_(row['Duration Seconds'], 0),
      category: String(row['Category'] || ALL_VALUE),
      subject: String(row['Subject'] || ALL_VALUE),
      topic: String(row['Topic'] || ALL_VALUE),
      difficulty: String(row['Difficulty'] || ALL_VALUE),
      feedbackMode: String(row['Feedback Mode'] || ''),
    };
  }).filter(function(item) { return item.studentKey.trim(); });
}

function analyticsNormalizeAnswers_(rows, questionMap) {
  return rows.map(function(row) {
    const questionId = String(row['QuestionID'] || '').trim();
    const question = questionMap[questionId] || {};
    const timestamp = analyticsDate_(row['Timestamp']);
    return {
      timestamp: analyticsFormatDate_(timestamp),
      timestampMs: timestamp ? timestamp.getTime() : 0,
      questionNumber: analyticsNumber_(row['QuestionNumber'], 0),
      questionId: questionId,
      questionText: question.questionText || '',
      category: String(row['Category'] || question.category || 'Unspecified'),
      subject: String(row['Subject'] || question.subject || 'Unspecified'),
      topic: String(row['Topic'] || question.topic || 'Unspecified'),
      difficulty: String(row['Difficulty'] || question.difficulty || 'Unspecified'),
      userAnswer: String(row['UserAnswer'] || '').trim(),
      correctAnswer: String(row['CorrectAnswer'] || question.correctAnswer || '').trim(),
      isCorrect: analyticsBool_(row['IsCorrect']),
      timedOut: analyticsBool_(row['TimedOut']),
      markedForReview: analyticsBool_(row['MarkedForReview']),
      timeSpentSeconds: analyticsNumber_(row['TimeSpentSeconds'], 0),
    };
  }).filter(function(item) { return item.questionId; });
}

function analyticsKpis_(attempts, answers, bank) {
  const correctAnswers = answers.filter(function(answer) { return answer.isCorrect; }).length;
  const timedOut = answers.filter(function(answer) { return answer.timedOut; }).length;
  const marked = answers.filter(function(answer) { return answer.markedForReview; }).length;
  const durationValues = attempts.map(function(attempt) { return attempt.durationSeconds; }).filter(function(value) { return value > 0; });
  return {
    attempts: attempts.length,
    uniqueStudents: analyticsUnique_(attempts.map(function(attempt) { return attempt.studentKey; })),
    totalAnswers: answers.length,
    avgAttemptPercentage: attempts.length ? analyticsRound_(analyticsAverage_(attempts.map(function(attempt) { return attempt.percentage; }))) : 0,
    answerAccuracy: answers.length ? analyticsRound_(correctAnswers / answers.length * 100) : 0,
    avgDurationSeconds: durationValues.length ? analyticsRound_(analyticsAverage_(durationValues)) : 0,
    timeoutRate: answers.length ? analyticsRound_(timedOut / answers.length * 100) : 0,
    markedForReviewRate: answers.length ? analyticsRound_(marked / answers.length * 100) : 0,
    validQuestions: bank.questions.length,
    invalidQuestionRows: bank.invalidQuestions,
  };
}

function analyticsAnswerGroups_(answers, key) {
  const groups = {};
  answers.forEach(function(answer) {
    const label = String(answer[key] || 'Unspecified').trim() || 'Unspecified';
    if (!groups[label]) groups[label] = { label: label, answers: 0, correct: 0, timedOut: 0, totalTime: 0, timedAnswers: 0 };
    groups[label].answers++;
    if (answer.isCorrect) groups[label].correct++;
    if (answer.timedOut) groups[label].timedOut++;
    if (answer.timeSpentSeconds > 0) {
      groups[label].totalTime += answer.timeSpentSeconds;
      groups[label].timedAnswers++;
    }
  });

  return Object.keys(groups).map(function(label) {
    const group = groups[label];
    return {
      label: label,
      answers: group.answers,
      correct: group.correct,
      accuracy: group.answers ? analyticsRound_(group.correct / group.answers * 100) : 0,
      timeoutRate: group.answers ? analyticsRound_(group.timedOut / group.answers * 100) : 0,
      avgTimeSeconds: group.timedAnswers ? analyticsRound_(group.totalTime / group.timedAnswers) : 0,
    };
  }).sort(function(a, b) { return b.answers - a.answers || a.label.localeCompare(b.label); });
}

function analyticsMissedQuestions_(answers) {
  const groups = {};
  answers.forEach(function(answer) {
    if (!groups[answer.questionId]) {
      groups[answer.questionId] = {
        questionId: answer.questionId,
        questionText: answer.questionText,
        category: answer.category,
        topic: answer.topic,
        difficulty: answer.difficulty,
        attempts: 0,
        misses: 0,
        timedOut: 0,
        correctAnswer: answer.correctAnswer,
      };
    }
    groups[answer.questionId].attempts++;
    if (!answer.isCorrect) groups[answer.questionId].misses++;
    if (answer.timedOut) groups[answer.questionId].timedOut++;
  });

  return Object.keys(groups).map(function(questionId) {
    const group = groups[questionId];
    group.missRate = group.attempts ? analyticsRound_(group.misses / group.attempts * 100) : 0;
    group.timeoutRate = group.attempts ? analyticsRound_(group.timedOut / group.attempts * 100) : 0;
    return group;
  }).filter(function(group) { return group.misses > 0; }).sort(function(a, b) {
    return b.missRate - a.missRate || b.misses - a.misses;
  });
}

function analyticsStudents_(attempts) {
  const groups = {};
  attempts.forEach(function(attempt) {
    const key = attempt.studentKey;
    if (!groups[key]) groups[key] = { student: attempt.student, attempts: 0, best: 0, total: 0, latest: '', latestMs: 0 };
    groups[key].attempts++;
    groups[key].best = Math.max(groups[key].best, attempt.percentage);
    groups[key].total += attempt.percentage;
    if (attempt.timestampMs >= groups[key].latestMs) {
      groups[key].latestMs = attempt.timestampMs;
      groups[key].latest = attempt.timestamp;
    }
  });

  return Object.keys(groups).map(function(key) {
    const group = groups[key];
    return {
      student: group.student,
      attempts: group.attempts,
      bestPercentage: analyticsRound_(group.best),
      averagePercentage: analyticsRound_(group.total / group.attempts),
      latestAttempt: group.latest,
    };
  }).sort(function(a, b) { return b.averagePercentage - a.averagePercentage; });
}

function analyticsQuestionMap_(questions) {
  const map = {};
  questions.forEach(function(question) { map[question.questionId] = question; });
  return map;
}

function analyticsColumnMap_(headers) {
  const map = {};
  headers.forEach(function(header, index) { if (header) map[header] = index; });
  return map;
}

function analyticsCell_(row, colMap, columnName) {
  const index = colMap[columnName];
  if (index === undefined) return '';
  return String(row[index] === null || row[index] === undefined ? '' : row[index]).trim();
}

function analyticsNormalizeFilters_(filters) {
  return {
    category: String(filters.category || ALL_VALUE).trim() || ALL_VALUE,
    subject: String(filters.subject || ALL_VALUE).trim() || ALL_VALUE,
    topic: String(filters.topic || ALL_VALUE).trim() || ALL_VALUE,
    difficulty: String(filters.difficulty || ALL_VALUE).trim() || ALL_VALUE,
  };
}

function analyticsMatches_(actual, expected) {
  return !expected || expected === ALL_VALUE || String(actual || '').trim() === expected;
}

function analyticsCountBy_(items, key) {
  const counts = {};
  (items || []).forEach(function(item) {
    const label = String(item[key] || 'Unspecified').trim() || 'Unspecified';
    counts[label] = (counts[label] || 0) + 1;
  });
  return counts;
}

function analyticsScore_(scoreValue, totalValue) {
  if (typeof scoreValue === 'string' && scoreValue.indexOf('/') !== -1) {
    const parts = scoreValue.split('/');
    return { score: Number(parts[0]) || 0, total: Number(parts[1]) || Number(totalValue) || 0 };
  }
  return { score: Number(scoreValue) || 0, total: Number(totalValue) || 0 };
}

function analyticsBool_(value) {
  if (typeof value === 'boolean') return value;
  const text = String(value || '').trim().toLowerCase();
  return text === 'true' || text === 'yes' || text === '1';
}

function analyticsNumber_(value, fallback) {
  const numeric = Number(value);
  return isNaN(numeric) ? fallback : numeric;
}

function analyticsAverage_(values) {
  const cleanValues = values.filter(function(value) { return typeof value === 'number' && !isNaN(value); });
  return cleanValues.length ? cleanValues.reduce(function(sum, value) { return sum + value; }, 0) / cleanValues.length : 0;
}

function analyticsUnique_(values) {
  const seen = {};
  values.forEach(function(value) { if (value) seen[value] = true; });
  return Object.keys(seen).length;
}

function analyticsDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') return isNaN(value.getTime()) ? '' : value;
  const date = new Date(value);
  return isNaN(date.getTime()) ? '' : date;
}

function analyticsFormatDate_(value) {
  const date = analyticsDate_(value);
  return date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : String(value || '');
}

function analyticsMaskName_(name) {
  const parts = String(name || '').trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return 'Unknown';
  if (parts.length === 1) return parts[0].charAt(0) + '***';
  return parts[0].charAt(0) + '*** ' + parts[parts.length - 1].charAt(0) + '***';
}

function analyticsMaskId_(idNumber) {
  const value = String(idNumber || '').trim();
  if (value.length <= 4) return '****';
  return '***' + value.slice(-4);
}

function analyticsRound_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}
