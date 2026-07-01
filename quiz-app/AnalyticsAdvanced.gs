/**
 * Advanced read-only analytics helpers for the Civil Engineering Quiz App.
 *
 * This file adds deeper analytics without changing existing spreadsheet rows.
 * It can be copied into Apps Script together with Code.gs, Analytics.gs, and
 * QuizPage.html.
 *
 * @version 1.0.0
 * @license MIT
 */

function getAdvancedQuizAnalytics() {
  const base = getQuizAnalytics();
  if (!base || !base.success) return base;

  const ss = SpreadsheetApp.openById(analyticsSpreadsheetId_());
  const users = analyticsReadRows_(ss, USERS_SHEET_NAME);
  const responses = analyticsReadRows_(ss, RESPONSES_SHEET_NAME);
  const bank = analyticsQuestionBank_();
  const questionMap = analyticsQuestionMap_(bank.questions);
  const attempts = analyticsNormalizeAttempts_(users);
  const answers = analyticsNormalizeAnswers_(responses, questionMap);

  base.kpis.medianAttemptPercentage = advancedMedian_(attempts.map(function(item) { return item.percentage; }));
  base.kpis.avgQuestionsPerAttempt = attempts.length ? advancedRound_(advancedAverage_(attempts.map(function(item) { return item.totalQuestions; }))) : 0;
  base.weakAreas = advancedWeakAreas_(base.breakdowns || {}).slice(0, 15);
  base.scoreDistribution = advancedScoreDistribution_(attempts);
  base.dailyTrend = advancedDailyTrend_(attempts).slice(-30);
  base.optionErrorPatterns = advancedOptionErrorPatterns_(answers).slice(0, 15);
  base.topMissedQuestions = advancedMissedQuestions_(answers).slice(0, 15);
  base.studentPerformance = advancedStudentPerformance_(attempts).slice(0, 25);

  return base;
}

function advancedWeakAreas_(breakdowns) {
  const rows = [];
  ['byCategory', 'bySubject', 'byTopic', 'byDifficulty'].forEach(function(groupKey) {
    const type = groupKey.replace('by', '');
    (breakdowns[groupKey] || []).forEach(function(row) {
      const accuracy = Number(row.accuracy) || 0;
      const timeoutRate = Number(row.timeoutRate) || 0;
      const answers = Number(row.answers) || 0;
      if (answers < 3) return;
      if (accuracy >= 75 && timeoutRate < 20) return;
      rows.push({
        type: type,
        label: row.label,
        answers: answers,
        accuracy: accuracy,
        timeoutRate: timeoutRate,
        avgTimeSeconds: Number(row.avgTimeSeconds) || 0,
        priorityScore: advancedRound_((100 - accuracy) + timeoutRate + Math.log(answers + 1) * 5),
      });
    });
  });
  return rows.sort(function(a, b) { return b.priorityScore - a.priorityScore || b.answers - a.answers; });
}

function advancedScoreDistribution_(attempts) {
  const buckets = [
    { label: '0-49%', min: 0, max: 49.999, count: 0 },
    { label: '50-69%', min: 50, max: 69.999, count: 0 },
    { label: '70-84%', min: 70, max: 84.999, count: 0 },
    { label: '85-100%', min: 85, max: 100, count: 0 },
  ];
  attempts.forEach(function(attempt) {
    const percentage = Number(attempt.percentage) || 0;
    const bucket = buckets.filter(function(item) { return percentage >= item.min && percentage <= item.max; })[0] || buckets[0];
    bucket.count++;
  });
  return buckets.map(function(bucket) {
    return {
      label: bucket.label,
      count: bucket.count,
      percentage: attempts.length ? advancedRound_(bucket.count / attempts.length * 100) : 0,
    };
  });
}

function advancedDailyTrend_(attempts) {
  const groups = {};
  attempts.forEach(function(attempt) {
    const key = attempt.timestamp ? String(attempt.timestamp).slice(0, 10) : 'Unknown Date';
    if (!groups[key]) groups[key] = { date: key, attempts: 0, totalPercentage: 0, totalDuration: 0, durationCount: 0 };
    groups[key].attempts++;
    groups[key].totalPercentage += Number(attempt.percentage) || 0;
    if (Number(attempt.durationSeconds) > 0) {
      groups[key].totalDuration += Number(attempt.durationSeconds);
      groups[key].durationCount++;
    }
  });
  return Object.keys(groups).sort().map(function(key) {
    const group = groups[key];
    return {
      date: group.date,
      attempts: group.attempts,
      avgPercentage: advancedRound_(group.totalPercentage / group.attempts),
      avgDurationSeconds: group.durationCount ? advancedRound_(group.totalDuration / group.durationCount) : 0,
    };
  });
}

function advancedOptionErrorPatterns_(answers) {
  const groups = {};
  answers.forEach(function(answer) {
    if (answer.isCorrect) return;
    const selected = String(answer.userAnswer || (answer.timedOut ? 'TIMEOUT' : 'BLANK')).toUpperCase();
    const key = answer.questionId + '|' + selected;
    if (!groups[key]) {
      groups[key] = {
        questionId: answer.questionId,
        selectedAnswer: selected,
        correctAnswer: answer.correctAnswer,
        topic: answer.topic,
        difficulty: answer.difficulty,
        count: 0,
      };
    }
    groups[key].count++;
  });
  return Object.keys(groups).map(function(key) { return groups[key]; }).sort(function(a, b) { return b.count - a.count; });
}

function advancedMissedQuestions_(answers) {
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
        totalTime: 0,
        timedAnswers: 0,
        correctAnswer: answer.correctAnswer,
        commonWrongAnswer: '',
        wrongAnswers: {},
      };
    }
    const group = groups[answer.questionId];
    group.attempts++;
    if (!answer.isCorrect) {
      group.misses++;
      const wrong = String(answer.userAnswer || (answer.timedOut ? 'TIMEOUT' : 'BLANK')).toUpperCase();
      group.wrongAnswers[wrong] = (group.wrongAnswers[wrong] || 0) + 1;
    }
    if (answer.timedOut) group.timedOut++;
    if (Number(answer.timeSpentSeconds) > 0) {
      group.totalTime += Number(answer.timeSpentSeconds);
      group.timedAnswers++;
    }
  });
  return Object.keys(groups).map(function(questionId) {
    const group = groups[questionId];
    group.missRate = group.attempts ? advancedRound_(group.misses / group.attempts * 100) : 0;
    group.timeoutRate = group.attempts ? advancedRound_(group.timedOut / group.attempts * 100) : 0;
    group.avgTimeSeconds = group.timedAnswers ? advancedRound_(group.totalTime / group.timedAnswers) : 0;
    group.commonWrongAnswer = advancedTopKey_(group.wrongAnswers);
    return group;
  }).filter(function(group) { return group.misses > 0; }).sort(function(a, b) {
    return b.missRate - a.missRate || b.misses - a.misses || b.timeoutRate - a.timeoutRate;
  });
}

function advancedStudentPerformance_(attempts) {
  const groups = {};
  attempts.sort(function(a, b) { return a.timestampMs - b.timestampMs; }).forEach(function(attempt) {
    const key = attempt.studentKey;
    if (!groups[key]) groups[key] = { student: attempt.student, attempts: 0, best: 0, total: 0, first: null, latest: null };
    groups[key].attempts++;
    groups[key].best = Math.max(groups[key].best, Number(attempt.percentage) || 0);
    groups[key].total += Number(attempt.percentage) || 0;
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
      bestPercentage: advancedRound_(group.best),
      averagePercentage: advancedRound_(group.total / group.attempts),
      latestPercentage: advancedRound_(Number(latest.percentage) || 0),
      improvement: advancedRound_((Number(latest.percentage) || 0) - (Number(first.percentage) || 0)),
      latestAttempt: latest.timestamp || '',
    };
  }).sort(function(a, b) { return b.averagePercentage - a.averagePercentage; });
}

function advancedAverage_(values) {
  const clean = values.filter(function(value) { return typeof value === 'number' && !isNaN(value); });
  return clean.length ? clean.reduce(function(sum, value) { return sum + value; }, 0) / clean.length : 0;
}

function advancedMedian_(values) {
  const clean = values.filter(function(value) { return typeof value === 'number' && !isNaN(value); }).sort(function(a, b) { return a - b; });
  if (!clean.length) return 0;
  const mid = Math.floor(clean.length / 2);
  return advancedRound_(clean.length % 2 ? clean[mid] : (clean[mid - 1] + clean[mid]) / 2);
}

function advancedTopKey_(counts) {
  const keys = Object.keys(counts || {});
  if (!keys.length) return '';
  keys.sort(function(a, b) { return counts[b] - counts[a] || a.localeCompare(b); });
  return keys[0] + ' (' + counts[keys[0]] + ')';
}

function advancedRound_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}
