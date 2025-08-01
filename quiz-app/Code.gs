/**
 * @fileoverview Backend script for a quiz web application using Google Sheets.
 * Handles fetching questions, processing answers, and recording results.
 *
 * @version 3.0
 * @author Your Name
 * @license MIT
 */

// --- Configuration ---
const SPREADSHEET_ID = "1qQw7B6sRrTkbGPViBmYqqaJEqwQr-7P0jeNuPirrgpY"; // IMPORTANT: Replace with your Google Sheet ID
const QUESTIONS_SHEET_NAME = "Questions";
const RESPONSES_SHEET_NAME = "Responses";
const USERS_SHEET_NAME = "Users";

// --- Column Name Constants for "Questions" Sheet (case-sensitive) ---
const Q_COL_ID = "Question ID";
const Q_COL_CATEGORY = "Category";
const Q_COL_SUBJECT = "Subject";
const Q_COL_TOPIC = "Topic";
const Q_COL_TEXT = "Question Text";
const Q_COL_OPTION_A = "OptionA";
const Q_COL_OPTION_B = "OptionB";
const Q_COL_OPTION_C = "OptionC";
const Q_COL_OPTION_D = "OptionD";
const Q_COL_ANSWER = "Answer"; // Correct option key (A, B, C, or D)
const Q_COL_EXPLANATION = "Explanation";
const Q_COL_IMAGE_URL = "ImageURL";

// --- Column Name Constants for "Responses" Sheet ---
const R_COL_TIMESTAMP = "Timestamp";
const R_COL_USER_ID = "UserID";
const R_COL_QUIZ_ID = "QuizIdentifier";
const R_COL_QUESTION_ID = "QuestionID";
const R_COL_USER_ANSWER = "UserAnswer";
const R_COL_IS_CORRECT = "IsCorrect";
const R_COL_SCORE = "ScoreAwarded";
const R_COL_TIMED_OUT = "TimedOut";

// --- Column Name Constants for "Users" Sheet ---
const U_COL_TIMESTAMP = "Timestamp";
const U_COL_NAME = "Name";
const U_COL_ID_NUMBER = "ID Number";
const U_COL_TOTAL_SCORE = "Total Score";
const U_COL_MODE = "Mode";

// --- Web App Entry Point ---
function doGet(e) {
  let htmlOutput = HtmlService.createHtmlOutputFromFile('QuizPage')
    .setTitle('Civil Service Examination Reviewer')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return htmlOutput;
}

// --- Utility Functions ---

/**
 * Creates a map of header names to their column index (0-based).
 * @param {string[]} headers - The array of header strings.
 * @returns {Object} A map where keys are header names and values are column indices.
 */
function getColumnMap(headers) {
  const map = {};
  headers.forEach((header, index) => {
    map[String(header).trim()] = index;
  });
  return map;
}

/**
 * Ensures a sheet exists and its headers are set correctly.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The spreadsheet object.
 * @param {string} sheetName - The name of the sheet to check/create.
 * @param {string[]} expectedHeaders - An array of expected header names.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} The sheet object, or null on failure.
 */
function ensureSheet(ss, sheetName, expectedHeaders) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`Sheet "${sheetName}" created.`);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(expectedHeaders);
    sheet.setFrozenRows(1);
    Logger.log(`Headers added to "${sheetName}" sheet.`);
    SpreadsheetApp.flush();
  }
  return sheet;
}


// --- Quiz Data Functions ---

/**
 * Gets all required data for the quiz setup and execution in one call.
 * Fetches filter options and all questions matching the filters.
 * @param {Object} filters - User-selected filters for the quiz.
 * @returns {Object} A result object with success status, data, or an error message.
 */
function getInitialQuizData(filters) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(QUESTIONS_SHEET_NAME);
    if (!sheet) {
      return { success: false, error: `Sheet "${QUESTIONS_SHEET_NAME}" not found.` };
    }

    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() <= 1) {
      return { success: false, error: `No questions found in "${QUESTIONS_SHEET_NAME}".` };
    }

    const data = dataRange.getValues();
    const headers = data[0];
    const colMap = getColumnMap(headers);
    
    // Validate that all essential columns exist
    const requiredCols = [Q_COL_ID, Q_COL_TEXT, Q_COL_OPTION_A, Q_COL_OPTION_B, Q_COL_ANSWER, Q_COL_EXPLANATION];
    for (const col of requiredCols) {
      if (colMap[col] === undefined) {
        return { success: false, error: `Required column "${col}" not found in "${QUESTIONS_SHEET_NAME}".` };
      }
    }
    
    const questionsData = data.slice(1);

    // --- 1. Get Filter Options for the UI ---
    const getUniqueSorted = (index) => [...new Set(questionsData.map(row => String(row[index] || '').trim()).filter(Boolean))].sort();
    const filterOptions = {
        categories: getUniqueSorted(colMap[Q_COL_CATEGORY]),
        subjects: getUniqueSorted(colMap[Q_COL_SUBJECT]),
        topics: getUniqueSorted(colMap[Q_COL_TOPIC])
    };

    // --- 2. Filter Questions based on criteria ---
    let filteredRows = questionsData.filter(row => {
      if (!row[colMap[Q_COL_ID]]) return false; // Must have an ID
      let match = true;
      if (filters.category && filters.category !== "All" && String(row[colMap[Q_COL_CATEGORY]] || '').trim() !== filters.category) match = false;
      if (filters.subject && filters.subject !== "All" && String(row[colMap[Q_COL_SUBJECT]] || '').trim() !== filters.subject) match = false;
      if (filters.topic && filters.topic !== "All" && String(row[colMap[Q_COL_TOPIC]] || '').trim() !== filters.topic) match = false;
      return match;
    });

    if (filteredRows.length === 0) {
      return { success: false, error: "No questions match the selected filters." };
    }

    // Randomize (Fisher-Yates Shuffle)
    if (filters.randomize) {
      for (let i = filteredRows.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [filteredRows[i], filteredRows[j]] = [filteredRows[j], filteredRows[i]];
      }
    }

    // Slice to the requested number of questions
    const requestedCount = parseInt(filters.count, 10) || 10;
    if (filteredRows.length > requestedCount) {
      filteredRows = filteredRows.slice(0, requestedCount);
    }
    
    // Map to client-side objects, NOW INCLUDING answer and explanation
    const questionObjects = filteredRows.map(row => ({
      questionId: String(row[colMap[Q_COL_ID]] || '').trim(),
      category: String(row[colMap[Q_COL_CATEGORY]] || '').trim(),
      subject: String(row[colMap[Q_COL_SUBJECT]] || '').trim(),
      topic: String(row[colMap[Q_COL_TOPIC]] || '').trim(),
      questionText: String(row[colMap[Q_COL_TEXT]] || '').trim(),
      options: { // Send options as an object for easier mapping
          A: String(row[colMap[Q_COL_OPTION_A]] || '').trim(),
          B: String(row[colMap[Q_COL_OPTION_B]] || '').trim(),
          C: String(row[colMap[Q_COL_OPTION_C]] || '').trim(),
          D: String(row[colMap[Q_COL_OPTION_D]] || '').trim(),
      },
      correctAnswer: String(row[colMap[Q_COL_ANSWER]] || '').trim().toUpperCase(),
      explanation: String(row[colMap[Q_COL_EXPLANATION]] || "No explanation provided.").trim(),
      imageUrl: String(row[colMap[Q_COL_IMAGE_URL]] || '').trim()
    })).filter(q => q.questionId && q.questionText && q.options.A && q.options.B && q.correctAnswer); // Basic validation

    if (questionObjects.length === 0) {
      return { success: false, error: "No valid questions were available after filtering." };
    }

    return { 
        success: true, 
        questions: questionObjects,
        filterOptions: filterOptions // Also return the latest filter options
    };

  } catch (error) {
    Logger.log(`Error in getInitialQuizData: ${error.toString()}\nStack: ${error.stack}`);
    return { success: false, error: `Server error fetching data: ${error.message}` };
  }
}


/**
 * Records the final summary and all individual responses in a single batch operation.
 * @param {Object} data - An object containing all necessary result information from the client.
 * @returns {Object} A result object indicating success or failure.
 */
function recordFullQuizResults(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // Wait up to 20 seconds

  try {
    const { userDetails, quizSummary, responses } = data;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date();

    // --- 1. Record to "Users" sheet ---
    const usersHeaders = [U_COL_TIMESTAMP, U_COL_NAME, U_COL_ID_NUMBER, U_COL_TOTAL_SCORE, U_COL_MODE];
    const usersSheet = ensureSheet(ss, USERS_SHEET_NAME, usersHeaders);
    if (!usersSheet) {
      return { success: false, error: `Critical Error: Sheet "${USERS_SHEET_NAME}" could not be accessed.` };
    }
    
    const totalScoreString = `${quizSummary.score}/${quizSummary.totalQuestions}`;
    const modeString = `NumQ: ${quizSummary.mode.numQuestions}, Time: ${quizSummary.mode.timePerQuestion}, Cat: ${quizSummary.mode.category}, Sub: ${quizSummary.mode.subject}, Top: ${quizSummary.mode.topic}`;
    
    usersSheet.appendRow([
      timestamp,
      userDetails.name,
      userDetails.idNumber,
      totalScoreString,
      modeString
    ]);
    Logger.log(`User result summary recorded for ${userDetails.name}`);

    // --- 2. Record all individual responses to "Responses" sheet ---
    if (responses && responses.length > 0) {
        const responsesHeaders = [R_COL_TIMESTAMP, R_COL_USER_ID, R_COL_QUIZ_ID, R_COL_QUESTION_ID, R_COL_USER_ANSWER, R_COL_IS_CORRECT, R_COL_SCORE, R_COL_TIMED_OUT];
        const responsesSheet = ensureSheet(ss, RESPONSES_SHEET_NAME, responsesHeaders);
         if (!responsesSheet) {
            return { success: false, error: `Critical Error: Sheet "${RESPONSES_SHEET_NAME}" could not be accessed.` };
        }

        const quizIdentifier = `${quizSummary.mode.category}-${quizSummary.mode.subject}-${quizSummary.mode.topic}`;

        // Prepare a 2D array for batch writing
        const rowsToAdd = responses.map(res => [
            timestamp,
            userDetails.name, // Use name as UserID for consistency
            quizIdentifier,
            res.questionId,
            res.userAnswer,
            res.isCorrect,
            res.isCorrect ? 1 : 0,
            res.timedOut
        ]);

        // Use setValues for efficient batch append
        responsesSheet.getRange(responsesSheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
        Logger.log(`${responses.length} individual responses recorded for ${userDetails.name}`);
    }

    return { success: true, message: "Quiz results recorded successfully." };

  } catch (error) {
    Logger.log(`Error in recordFullQuizResults for ${data.userDetails.name}: ${error.toString()}\nStack: ${error.stack}`);
    return { success: false, error: `Server error recording quiz results: ${error.message}` };
  } finally {
    lock.releaseLock();
  }
}