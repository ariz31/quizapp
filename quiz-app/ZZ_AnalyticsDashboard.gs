/**
 * Dashboard endpoint override for the Civil Engineering Quiz App.
 *
 * Apps Script shares one global namespace across script files. This file is
 * intentionally named with a ZZ prefix so the dashboard endpoint can return
 * the richer analytics model while keeping the original Analytics.gs helpers
 * available.
 *
 * @version 1.0.0
 * @license MIT
 */

function getQuizAnalytics() {
  try {
    const ss = SpreadsheetApp.openById(analyticsSpreadsheetId_());
    const users = analyticsReadRows_(ss, USERS_SHEET_NAME);
    const responses = analyticsReadRows_(ss, RESPONSES_SHEET_NAME);
    const bank = analyticsQuestionBank_();
    const questionMap = analyticsQuestionMap_(bank.questions);
    const attempts = analyticsNormalizeAttempts_(users);
    const answers = analyticsNormalizeAnswers_(responses, questionMap);
    const breakdowns = {
      byCategory: analyticsAnswerGroups_(answers, 'category'),
      bySubject: analyticsAnswerGroups_(answers, 'subject'),
      byTopic: analyticsAnswerGroups_(answers, 'topic'),
      byDifficulty: analyticsAnswerGroups_(answers, 'difficulty'),
    };
    const kpis = analyticsKpis_(attempts, answers, bank);

    kpis.medianAttemptPercentage = advancedMedian_(attempts.map(function(item) { return item.percentage; }));
    kpis.avgQuestionsPerAttempt = attempts.length ? advancedRound_(advancedAverage_(attempts.map(function(item) { return item.totalQuestions; }))) : 0;

    return {
      success: true,
      generatedAt: analyticsFormatDate_(new Date()),
      kpis: kpis,
      breakdowns: breakdowns,
      topMissedQuestions: advancedMissedQuestions_(answers).slice(0, 15),
      weakAreas: advancedWeakAreas_(breakdowns).slice(0, 15),
      scoreDistribution: advancedScoreDistribution_(attempts),
      dailyTrend: advancedDailyTrend_(attempts).slice(-30),
      optionErrorPatterns: advancedOptionErrorPatterns_(answers).slice(0, 15),
      studentPerformance: advancedStudentPerformance_(attempts).slice(0, 25),
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
    Logger.log('getQuizAnalytics dashboard endpoint failed: ' + error.toString() + '\n' + error.stack);
    return { success: false, error: 'Unable to build analytics: ' + error.message };
  }
}
