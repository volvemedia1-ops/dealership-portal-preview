/**
 * Credit Builder Flow • Post-Meeting Portal URL & Gmail Draft Generator
 *
 * Reads newly created Google Meet "Notes by Gemini" Google Docs in Drive,
 * matches them against the "Booked Meetings" tab in Cold Call Broker List,
 * extracts structured meeting parameters via Gemini API (zero guessing),
 * builds the Option A1 personalised portal link, and saves a Gmail Draft
 * (never auto-sent) while logging status back to the sheet non-destructively.
 *
 * Setup:
 *   1. Open Cold Call Broker List -> Extensions -> Apps Script.
 *   2. Paste FollowUpDraftGenerator.gs and enable appsscript.json manifest.
 *   3. Set Script Property GEMINI_API_KEY (Project Settings -> Script Properties).
 *   4. Run testDraftGeneratorDryRun() to verify formatting in logs.
 *   5. Run installTenMinuteTrigger() once to automate every 10 minutes.
 */

var CBF_CONFIG = {
  SPREADSHEET_ID: '1BkZafvLWWWrKm56r_c8IThjIEXcPOZHWnEF4unA-K5g',
  SHEET_NAME: 'Booked Meetings',
  DEFAULT_PORTAL_BASE_URL: 'https://dealerships.volvemedia.com/',
  DEFAULT_GEMINI_MODEL: 'gemini-2.5-flash',
  DEFAULT_LOOKBACK_DAYS: 7,
  MAX_TRANSCRIPT_CHARS: 24000,
  EXISTING_CORE_COL_COUNT: 25, // Columns A-Y must never be overwritten
  TRACKING_HEADERS: [
    'Follow-Up Status',
    'Portal URL',
    'Gmail Draft ID',
    'Notes Doc ID',
    'Extracted JSON',
    'Processed At'
  ],
  VALID_CONCERNS: {
    exclusivity: true,
    popia: true,
    staff: true,
    cost: true,
    commitment: true
  },
  SENDER_SIGNATURE: {
    name: 'Kea Letsoalo',
    company: 'Volve Media | Credit Builder Flow',
    contact: '+27 79 476 2500 | volve@volvemedia.com'
  }
};

/**
 * Main automated entry point (intended for 10-minute time-driven trigger).
 * Reads new Gemini Notes Docs, matches to Booked Meetings rows, calls Gemini,
 * creates Gmail Drafts (never sends), and logs results after Column Y.
 */
function processCompletedMeetings() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    Logger.log('[CBF] Skipped run: another execution already holds the lock.');
    return;
  }

  try {
    var props = PropertiesService.getScriptProperties();
    var apiKey = (props.getProperty('GEMINI_API_KEY') || '').trim();
    if (!apiKey) {
      throw new Error('Missing Script Property GEMINI_API_KEY. Configure it under Project Settings -> Script Properties.');
    }

    var portalBaseUrl = (props.getProperty('PORTAL_BASE_URL') || CBF_CONFIG.DEFAULT_PORTAL_BASE_URL).trim();
    var geminiModel = (props.getProperty('GEMINI_MODEL') || CBF_CONFIG.DEFAULT_GEMINI_MODEL).trim();
    var lookbackDays = parseInt(props.getProperty('LOOKBACK_DAYS') || String(CBF_CONFIG.DEFAULT_LOOKBACK_DAYS), 10);
    if (isNaN(lookbackDays) || lookbackDays < 1) {
      lookbackDays = CBF_CONFIG.DEFAULT_LOOKBACK_DAYS;
    }

    var ss = SpreadsheetApp.openById(CBF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CBF_CONFIG.SHEET_NAME);
    if (!sheet) {
      throw new Error('Sheet "' + CBF_CONFIG.SHEET_NAME + '" not found in spreadsheet ' + CBF_CONFIG.SPREADSHEET_ID);
    }

    var colMap = ensureTrackingColumns_(sheet);
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) {
      Logger.log('[CBF] No rows in "' + CBF_CONFIG.SHEET_NAME + '".');
      return;
    }

    var allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var processedDocIds = collectProcessedDocIds_(allValues, colMap);
    var candidates = findCandidateNotesDocs_(lookbackDays, processedDocIds);

    Logger.log('[CBF] Found ' + candidates.length + ' candidate unprocessed Gemini Notes document(s).');

    for (var i = 0; i < candidates.length; i++) {
      var fileInfo = candidates[i];
      try {
        var docContent = readDocumentTabsAndLinks_(fileInfo.id);
        var matchResult = matchBookedMeetingRow_(fileInfo, docContent, allValues, colMap);

        if (!matchResult || matchResult.status === 'no_match') {
          Logger.log('[CBF] No Booked Meetings match for Doc "' + fileInfo.name + '" (' + fileInfo.id + ').');
          continue;
        }

        if (matchResult.status === 'ambiguous') {
          Logger.log('[CBF] Ambiguous match for Doc "' + fileInfo.name + '" across rows: ' + matchResult.candidateRows.join(', '));
          continue;
        }

        var rowIdx = matchResult.rowIndex; // 0-based index into allValues (sheet row = rowIdx + 1)
        var rowNum = rowIdx + 1;
        var rowObj = rowToObject_(allValues[rowIdx], colMap);

        if (String(rowObj.notesDocId || '').trim() || String(rowObj.gmailDraftId || '').trim()) {
          Logger.log('[CBF] Row ' + rowNum + ' (' + rowObj.company + ') already processed; skipping.');
          continue;
        }

        var extracted = extractMeetingDetailsWithGemini_(
          apiKey,
          geminiModel,
          rowObj.company,
          rowObj.firstName,
          docContent.combinedText
        );

        if (extracted.isDealerMeeting === false) {
          Logger.log('[CBF] Gemini flagged Doc "' + fileInfo.name + '" as non-dealer/internal session; skipping draft creation.');
          continue;
        }

        var missingDeclines = !(extracted.declines && extracted.declines > 0);
        var primaryEmail = validateEmailOrNull_(rowObj.email);
        var dpEmail = validateEmailOrNull_(extracted.dpEmail);
        var recipients = [];
        if (primaryEmail) recipients.push(primaryEmail);
        if (dpEmail && dpEmail.toLowerCase() !== (primaryEmail || '').toLowerCase()) {
          recipients.push(dpEmail);
        }

        var meetingDateShort = formatMeetingDateShort_(rowObj.meetingDate);
        var portalParams = {
          dealer: rowObj.company,
          contact: rowObj.firstName,
          dp: extracted.dpName || '',
          date: meetingDateShort,
          declines: extracted.declines,
          margin: extracted.marginK,
          banks: extracted.banks || '',
          concern: extracted.concern || ''
        };
        var portalUrl = buildPortalUrl_(portalBaseUrl, portalParams);

        var statusParts = [];
        var draftId = '';

        if (recipients.length === 0) {
          statusParts.push('Needs review (Missing Email)');
        } else {
          var emailContent = buildFollowUpEmail_({
            dealer: rowObj.company,
            contact: rowObj.firstName,
            dpName: extracted.dpName || '',
            hasDpEmail: Boolean(dpEmail),
            declines: extracted.declines,
            portalUrl: portalUrl
          });

          var draft = GmailApp.createDraft(
            recipients.join(', '),
            emailContent.subject,
            emailContent.plainTextBody,
            { htmlBody: emailContent.htmlBody }
          );
          draftId = draft.getId();
          statusParts.push(missingDeclines ? 'Needs review' : 'Draft Created');
        }

        if (missingDeclines && recipients.length > 0) {
          statusParts = ['Needs review'];
        } else if (missingDeclines && recipients.length === 0) {
          statusParts = ['Needs review (Missing Email & Declines)'];
        }

        var nowIso = Utilities.formatDate(new Date(), 'Africa/Johannesburg', 'yyyy-MM-dd HH:mm:ss z');
        var auditJson = JSON.stringify({
          dpName: extracted.dpName,
          dpEmail: extracted.dpEmail,
          declines: extracted.declines,
          marginK: extracted.marginK,
          banks: extracted.banks,
          concern: extracted.concern,
          matchedBy: matchResult.matchedBy
        });

        writeTrackingCells_(sheet, rowNum, colMap, {
          status: statusParts.join(' '),
          portalUrl: portalUrl,
          draftId: draftId,
          docId: fileInfo.id,
          extractedJson: auditJson,
          processedAt: nowIso
        });

        Logger.log(
          '[CBF] Processed Row ' + rowNum + ' (' + rowObj.company + ') -> Status: ' +
          statusParts.join(' ') + ', DraftID: ' + (draftId || 'NONE') + ', URL: ' + portalUrl
        );
      } catch (docErr) {
        Logger.log('[CBF] Error processing Doc "' + fileInfo.name + '" (' + fileInfo.id + '): ' + docErr.message);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Ensures tracking columns exist strictly AFTER Column Y (col >= 26),
 * preserving all existing A-Y columns untouched.
 */
function ensureTrackingColumns_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), CBF_CONFIG.EXISTING_CORE_COL_COUNT);
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerIndex = {};

  for (var c = 0; c < headerRow.length; c++) {
    var name = String(headerRow[c] || '').trim();
    if (name) {
      headerIndex[name] = c + 1; // 1-based column number
    }
  }

  var nextCol = lastCol + 1;
  for (var t = 0; t < CBF_CONFIG.TRACKING_HEADERS.length; t++) {
    var trackName = CBF_CONFIG.TRACKING_HEADERS[t];
    if (!headerIndex[trackName]) {
      sheet.getRange(1, nextCol).setValue(trackName);
      headerIndex[trackName] = nextCol;
      nextCol++;
    }
  }

  return headerIndex;
}

/**
 * Collects Notes Doc IDs already logged in the tracking column so we never duplicate drafts.
 */
function collectProcessedDocIds_(allValues, colMap) {
  var set = {};
  var docCol = colMap['Notes Doc ID'] ? colMap['Notes Doc ID'] - 1 : -1;
  if (docCol < 0) return set;

  for (var r = 1; r < allValues.length; r++) {
    var id = String(allValues[r][docCol] || '').trim();
    if (id) set[id] = true;
  }
  return set;
}

/**
 * Searches Google Drive for recent Google Docs generated by Gemini Notes.
 */
function findCandidateNotesDocs_(lookbackDays, processedDocIds) {
  var cutoff = new Date(Date.now() - lookbackDays * 24 * 60 * 60 * 1000);
  var cutoffStr = cutoff.toISOString().replace(/\.\d{3}Z$/, 'Z');
  var query =
    "mimeType = 'application/vnd.google-apps.document' " +
    "and trashed = false " +
    "and modifiedDate >= '" + cutoffStr + "' " +
    "and (title contains 'Notes by Gemini' or title contains 'Recover Lost Revenue')";

  var results = [];
  var files = DriveApp.searchFiles(query);
  while (files.hasNext()) {
    var file = files.next();
    var id = file.getId();
    if (processedDocIds[id]) continue;
    results.push({
      id: id,
      name: file.getName(),
      createdAt: file.getDateCreated(),
      updatedAt: file.getLastUpdated()
    });
  }

  results.sort(function (a, b) {
    return b.createdAt.getTime() - a.createdAt.getTime();
  });
  return results;
}

/**
 * Reads all tabs (Quick notes, Full notes, Transcript) from a Google Doc
 * and collects any calendar event ?eid= links embedded in Attachments.
 */
function readDocumentTabsAndLinks_(docId) {
  var doc = DocumentApp.openById(docId);
  var sections = [];
  var calendarEids = [];

  if (typeof doc.getTabs === 'function') {
    var tabs = doc.getTabs();
    collectTabsRecursively_(tabs, sections, calendarEids);
  }

  // Fallback for single-tab or legacy DocumentApp behavior
  if (sections.length === 0) {
    var legacyBody = doc.getBody();
    extractLinksFromContainer_(legacyBody, calendarEids);
    sections.push({
      title: 'Body',
      text: legacyBody.getText() || ''
    });
  }

  // Prioritize Quick notes & Full notes first, then Transcript up to budget
  sections.sort(function (a, b) {
    return tabPriority_(a.title) - tabPriority_(b.title);
  });

  var combinedParts = [];
  for (var i = 0; i < sections.length; i++) {
    var titleLower = String(sections[i].title || '').toLowerCase();
    var text = String(sections[i].text || '').trim();
    if (!text) continue;
    if (titleLower.indexOf('transcript') !== -1 && text.length > CBF_CONFIG.MAX_TRANSCRIPT_CHARS) {
      text = text.slice(0, CBF_CONFIG.MAX_TRANSCRIPT_CHARS) + '\n[...transcript truncated...]';
    }
    combinedParts.push('=== TAB: ' + sections[i].title + ' ===\n' + text);
  }

  return {
    combinedText: combinedParts.join('\n\n'),
    calendarEventIds: decodeCalendarEids_(calendarEids)
  };
}

function tabPriority_(title) {
  var t = String(title || '').toLowerCase();
  if (t.indexOf('full notes') !== -1) return 1;
  if (t.indexOf('quick notes') !== -1) return 2;
  if (t.indexOf('transcript') !== -1) return 4;
  return 3;
}

function collectTabsRecursively_(tabs, outSections, outCalendarEids) {
  if (!tabs) return;
  for (var i = 0; i < tabs.length; i++) {
    var tab = tabs[i];
    if (typeof tab.asDocumentTab === 'function') {
      try {
        var docTab = tab.asDocumentTab();
        var body = docTab.getBody();
        extractLinksFromContainer_(body, outCalendarEids);
        outSections.push({
          title: typeof tab.getTitle === 'function' ? tab.getTitle() : 'Tab',
          text: body.getText() || ''
        });
      } catch (e) {
        // Non-document tab or empty body
      }
    }
    if (typeof tab.getChildTabs === 'function') {
      collectTabsRecursively_(tab.getChildTabs(), outSections, outCalendarEids);
    }
  }
}

function extractLinksFromContainer_(container, outCalendarEids) {
  if (!container || typeof container.getNumChildren !== 'function') return;
  var num = container.getNumChildren();
  for (var i = 0; i < num; i++) {
    var child = container.getChild(i);
    var type = child.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH || type === DocumentApp.ElementType.LIST_ITEM) {
      var textEl = child.editAsText();
      var indices = textEl.getTextAttributeIndices();
      for (var j = 0; j < indices.length; j++) {
        var url = textEl.getLinkUrl(indices[j]);
        if (url && url.indexOf('calendar.google.com/calendar/event') !== -1) {
          outCalendarEids.push(url);
        }
      }
    }
  }
}

function decodeCalendarEids_(urls) {
  var ids = {};
  for (var i = 0; i < urls.length; i++) {
    var match = String(urls[i]).match(/[?&]eid=([^&#]+)/);
    if (!match) continue;
    try {
      var decodedBytes = Utilities.base64DecodeWebSafe(match[1]);
      var decodedStr = Utilities.newBlob(decodedBytes).getDataAsString('UTF-8');
      var eventIdPrefix = decodedStr.trim().split(/\s+/)[0];
      if (eventIdPrefix) {
        ids[eventIdPrefix.toLowerCase()] = true;
      }
    } catch (e) {
      // Ignore malformed base64 eid
    }
  }
  return Object.keys(ids);
}

/**
 * Matches a Notes Doc to a single Booked Meetings row:
 *   1) Exact Calendar Event ID match from Full notes Attachments hyperlink (?eid=).
 *   2) Dealer company name + meeting date match parsed from Doc title & text.
 */
function matchBookedMeetingRow_(fileInfo, docContent, allValues, colMap) {
  var rows = [];
  for (var r = 1; r < allValues.length; r++) {
    var obj = rowToObject_(allValues[r], colMap);
    if (!obj.company) continue;
    rows.push({ rowIndex: r, data: obj });
  }

  // Tier 1: Exact Calendar Event ID match
  if (docContent.calendarEventIds && docContent.calendarEventIds.length > 0) {
    var eidMatches = [];
    for (var i = 0; i < rows.length; i++) {
      var sheetEventRaw = String(rows[i].data.calendarEventId || '').trim().toLowerCase();
      var sheetEventPrefix = sheetEventRaw.split('@')[0];
      if (sheetEventPrefix && docContent.calendarEventIds.indexOf(sheetEventPrefix) !== -1) {
        eidMatches.push(rows[i]);
      }
    }
    if (eidMatches.length === 1) {
      return { status: 'matched', rowIndex: eidMatches[0].rowIndex, matchedBy: 'calendar_event_id' };
    }
    if (eidMatches.length > 1) {
      return {
        status: 'ambiguous',
        candidateRows: eidMatches.map(function (m) { return m.rowIndex + 1; })
      };
    }
  }

  // Tier 2: Company name in doc title / header + Meeting Date window
  var titleCompanyHint = parseCompanyHintFromTitle_(fileInfo.name);
  var docDateYmd = parseMeetingYmdFromDoc_(fileInfo.name, fileInfo.createdAt);
  var scored = [];

  for (var j = 0; j < rows.length; j++) {
    var row = rows[j];
    var company = String(row.data.company || '').trim();
    var normCompany = normalizeDealerName_(company);
    if (!normCompany || normCompany.length < 3) continue;

    var score = 0;
    if (titleCompanyHint && normalizeDealerName_(titleCompanyHint) === normCompany) {
      score += 100;
    } else if (titleCompanyHint && normalizeDealerName_(titleCompanyHint).indexOf(normCompany) !== -1) {
      score += 75;
    } else if (String(fileInfo.name || '').toLowerCase().indexOf(company.toLowerCase()) !== -1) {
      score += 70;
    } else if (normCompany.length >= 4 && normalizeDealerName_(fileInfo.name).indexOf(normCompany) !== -1) {
      score += 70;
    }

    if (score === 0) continue;

    // Require or boost date match within +/- 1 day
    var rowDateYmd = toYmdString_(row.data.meetingDate);
    if (docDateYmd && rowDateYmd) {
      var diffDays = Math.abs(ymdDiffDays_(docDateYmd, rowDateYmd));
      if (diffDays === 0) {
        score += 40;
      } else if (diffDays === 1) {
        score += 15;
      } else {
        score -= 60; // Different scheduled day
      }
    }

    var mStatus = String(row.data.meetingStatus || '').toLowerCase();
    if (mStatus === 'scheduled' || mStatus === 'completed') score += 10;
    if (mStatus.indexOf('reschedul') !== -1 || mStatus.indexOf('cancel') !== -1) score -= 30;

    if (score >= 70) {
      scored.push({ rowIndex: row.rowIndex, score: score });
    }
  }

  if (scored.length === 0) return { status: 'no_match' };

  scored.sort(function (a, b) { return b.score - a.score; });
  if (scored.length > 1 && scored[0].score === scored[1].score) {
    var tied = scored
      .filter(function (s) { return s.score === scored[0].score; })
      .map(function (s) { return s.rowIndex + 1; });
    return { status: 'ambiguous', candidateRows: tied };
  }

  return { status: 'matched', rowIndex: scored[0].rowIndex, matchedBy: 'title_and_date' };
}

function rowToObject_(rowArr, colMap) {
  function getVal(header) {
    var col = colMap[header];
    return col ? rowArr[col - 1] : '';
  }
  return {
    company: String(getVal('Company') || '').trim(),
    firstName: String(getVal('First Name') || '').trim(),
    title: String(getVal('Title') || '').trim(),
    email: String(getVal('Email') || '').trim(),
    contactNo: String(getVal('Contact no.') || '').trim(),
    province: String(getVal('Province') || '').trim(),
    meetingDate: getVal('Meeting Date'),
    meetingTime: getVal('Meeting Time'),
    meetingStatus: String(getVal('Meeting Status') || '').trim(),
    calendarEventId: String(getVal('Calendar Event ID') || '').trim(),
    callNotes: String(getVal('Call Notes') || '').trim(),
    followUpStatus: String(getVal('Follow-Up Status') || '').trim(),
    gmailDraftId: String(getVal('Gmail Draft ID') || '').trim(),
    notesDocId: String(getVal('Notes Doc ID') || '').trim()
  };
}

function parseCompanyHintFromTitle_(docTitle) {
  var s = String(docTitle || '');
  var m = s.match(/^(?:Recover\s+Lost\s+Revenue:\s*)?(.+?)\s*(?:x|&|\/|-)\s*Credit\s+Builder\s+Flow/i);
  if (m && m[1]) return m[1].trim();
  var m2 = s.match(/^(.+?)\s*-\s*\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}\s+UTC\s*-\s*Notes\s+by\s+Gemini/i);
  if (m2 && m2[1]) return m2[1].trim();
  return '';
}

function parseMeetingYmdFromDoc_(docTitle, fallbackCreatedDate) {
  var m = String(docTitle || '').match(/(\d{4})\/(\d{2})\/(\d{2})\s+(\d{2}):(\d{2})\s*UTC/i);
  if (m) {
    var utcDate = new Date(Date.UTC(
      parseInt(m[1], 10),
      parseInt(m[2], 10) - 1,
      parseInt(m[3], 10),
      parseInt(m[4], 10),
      parseInt(m[5], 10)
    ));
    return Utilities.formatDate(utcDate, 'Africa/Johannesburg', 'yyyy-MM-dd');
  }
  if (fallbackCreatedDate instanceof Date) {
    return Utilities.formatDate(fallbackCreatedDate, 'Africa/Johannesburg', 'yyyy-MM-dd');
  }
  return '';
}

function toYmdString_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Africa/Johannesburg', 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  var dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (dmy) {
    var dd = ('0' + dmy[1]).slice(-2);
    var mm = ('0' + dmy[2]).slice(-2);
    return dmy[3] + '-' + mm + '-' + dd;
  }
  return '';
}

function ymdDiffDays_(ymdA, ymdB) {
  var a = ymdA.split('-');
  var b = ymdB.split('-');
  var da = Date.UTC(parseInt(a[0], 10), parseInt(a[1], 10) - 1, parseInt(a[2], 10));
  var db = Date.UTC(parseInt(b[0], 10), parseInt(b[1], 10) - 1, parseInt(b[2], 10));
  return Math.round((da - db) / (24 * 60 * 60 * 1000));
}

function normalizeDealerName_(name) {
  return String(name || '')
    .toLowerCase()
    .replace(/\b(pty|ltd|cc|motors|auto|group|dealership|car|cars)\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

/**
 * Calls Gemini REST API with strict JSON schema guidance and zero guessing.
 */
function extractMeetingDetailsWithGemini_(apiKey, model, dealerName, contactName, notesText) {
  var endpoint = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(model) + ':generateContent';

  var prompt = [
    'You are extracting structured B2B meeting facts for an automotive dealership follow-up portal.',
    'Dealership: ' + dealerName,
    'Primary Contact (F&I): ' + (contactName || 'Unknown'),
    '',
    'CRITICAL RULES:',
    '1. Extract ONLY facts explicitly stated in the meeting notes/transcript below.',
    '2. NEVER guess or invent numbers. If a value is not explicitly stated, return null.',
    '3. Return strictly valid JSON with keys:',
    '   - isDealerMeeting: boolean (true if an actual meeting with "' + dealerName + '" happened; false if it was internal prep/cold calling or unrelated)',
    '   - dpName: string or null (Dealer Principal first name if mentioned)',
    '   - dpEmail: string or null (Dealer Principal email if explicitly mentioned)',
    '   - declines: integer or null (monthly finance declines mentioned at the dealership, e.g. 30)',
    '   - marginK: number or null (average front-end gross profit per car in ZAR thousands, e.g. R18,000 -> 18, R15k -> 15)',
    '   - banks: string or null (comma-separated SA finance banks mentioned, e.g. "WesBank,MFC")',
    '   - concern: one of "exclusivity", "popia", "staff", "cost", "commitment", or null (pick ONLY if a matching concern/objection was clearly raised)',
    '',
    'MEETING NOTES & TRANSCRIPT:',
    notesText
  ].join('\n');

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json'
    }
  };

  var res = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  var bodyText = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Gemini API HTTP ' + code + ': ' + bodyText.slice(0, 400));
  }

  var parsedResponse = JSON.parse(bodyText);
  var rawText = '';
  if (
    parsedResponse &&
    parsedResponse.candidates &&
    parsedResponse.candidates[0] &&
    parsedResponse.candidates[0].content &&
    parsedResponse.candidates[0].content.parts &&
    parsedResponse.candidates[0].content.parts[0]
  ) {
    rawText = parsedResponse.candidates[0].content.parts[0].text || '{}';
  }

  // Strip optional markdown fences if returned
  rawText = rawText.replace(/^```json\s*/i, '').replace(/```$/i, '').trim();
  var data = JSON.parse(rawText);

  return sanitizeExtractedFields_(data);
}

function sanitizeExtractedFields_(raw) {
  var obj = raw || {};

  var declinesNum = parseInt(obj.declines, 10);
  var declines = (!isNaN(declinesNum) && declinesNum > 0 && declinesNum < 2000) ? declinesNum : null;

  var marginNum = parseFloat(obj.marginK);
  var marginK = null;
  if (!isNaN(marginNum) && marginNum > 0) {
    marginK = marginNum >= 1000 ? Math.round((marginNum / 1000) * 10) / 10 : Math.round(marginNum * 10) / 10;
  }

  var banksStr = '';
  if (Array.isArray(obj.banks)) {
    banksStr = obj.banks.map(function (b) { return String(b || '').trim(); }).filter(Boolean).join(',');
  } else if (typeof obj.banks === 'string') {
    banksStr = obj.banks.split(',').map(function (b) { return b.trim(); }).filter(Boolean).join(',');
  }

  var concernKey = String(obj.concern || '').trim().toLowerCase();
  var concern = CBF_CONFIG.VALID_CONCERNS[concernKey] ? concernKey : null;

  var dpName = obj.dpName ? String(obj.dpName).trim() : null;
  if (dpName && /^(null|none|unknown|n\/a)$/i.test(dpName)) dpName = null;

  return {
    isDealerMeeting: obj.isDealerMeeting !== false,
    dpName: dpName,
    dpEmail: validateEmailOrNull_(obj.dpEmail),
    declines: declines,
    marginK: marginK,
    banks: banksStr || null,
    concern: concern
  };
}

function validateEmailOrNull_(val) {
  var s = String(val || '').trim();
  if (!s) return null;
  var m = s.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/);
  return m ? s : null;
}

/**
 * Formats sheet Meeting Date into concise Option A1 display format, e.g. "19 Nov".
 */
function formatMeetingDateShort_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Africa/Johannesburg', 'd MMM');
  }
  var s = String(val).trim();
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    var d = new Date(Date.UTC(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10), 12, 0, 0));
    return Utilities.formatDate(d, 'Africa/Johannesburg', 'd MMM');
  }
  return s;
}

/**
 * Builds the clean Option A1 personalised portal URL.
 * Omits null/empty optional parameters cleanly.
 */
function buildPortalUrl_(baseUrl, p) {
  var base = String(baseUrl || CBF_CONFIG.DEFAULT_PORTAL_BASE_URL).split('?')[0];
  var pairs = [];
  function pushParam(k, v) {
    if (v === null || v === undefined) return;
    var str = String(v).trim();
    if (!str) return;
    pairs.push(encodeURIComponent(k) + '=' + encodeURIComponent(str));
  }

  pushParam('dealer', p.dealer);
  pushParam('contact', p.contact);
  pushParam('dp', p.dp);
  pushParam('date', p.date);
  pushParam('declines', p.declines);
  pushParam('margin', p.margin);
  pushParam('banks', p.banks);
  pushParam('concern', p.concern);

  return pairs.length > 0 ? base + '?' + pairs.join('&') : base;
}

/**
 * Builds exact Option A1 Gmail Draft Subject, Plain Text Body, and HTML Body.
 */
function buildFollowUpEmail_(opts) {
  var dealer = String(opts.dealer || 'Your Dealership').trim();
  var contact = String(opts.contact || '').trim();
  var dpName = String(opts.dpName || '').trim();
  var hasDpEmail = Boolean(opts.hasDpEmail);
  var declines = opts.declines;
  var portalUrl = opts.portalUrl;

  var subject = 'Meeting Follow-Up: Credit Builder Flow x ' + dealer;

  var greeting = 'Hi ';
  if (contact && dpName && hasDpEmail && dpName.toLowerCase() !== contact.toLowerCase()) {
    greeting += contact + ' and ' + dpName + ',';
  } else if (contact) {
    greeting += contact + ',';
  } else if (dpName) {
    greeting += dpName + ',';
  } else {
    greeting += dealer + ' Team,';
  }

  var declineBullet = (declines && declines > 0)
    ? 'A summary of our discussion, including the ~' + declines + ' monthly declines mentioned'
    : 'A summary of our discussion and your dealership finance recovery opportunity';

  var forwardLine = hasDpEmail
    ? ''
    : 'Please feel free to forward this email to your Dealer Principal.';

  var plainLines = [
    greeting,
    '',
    'Thank you for your time today. As discussed, we have prepared a review portal for ' + dealer + ' that summarises our conversation, the commercial terms and the Master Service Agreement.',
    '',
    '[ Open the ' + dealer + ' Review Portal ]',
    portalUrl,
    '',
    'The portal includes:',
    '- ' + declineBullet,
    '- The fee structure: R0 retainer; R6,000 or 5% (incl. VAT), payable only per delivered vehicle',
    '- Case studies and the compliance framework (POPIA and exclusivity)',
    '- The Master Service Agreement, which can be reviewed and signed online',
    '',
    'Once the agreement is signed, your team will be directed to register on the dealership portal, where your F&I referral links are issued.'
  ];

  if (forwardLine) {
    plainLines.push('');
    plainLines.push(forwardLine);
  }

  plainLines.push('');
  plainLines.push('Kind regards,');
  plainLines.push(CBF_CONFIG.SENDER_SIGNATURE.name);
  plainLines.push(CBF_CONFIG.SENDER_SIGNATURE.company);
  plainLines.push(CBF_CONFIG.SENDER_SIGNATURE.contact);

  var esc = escapeHtml_;
  var htmlParts = [
    '<div style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #1f2937; line-height: 1.6;">',
    '<p>' + esc(greeting) + '</p>',
    '<p>Thank you for your time today. As discussed, we have prepared a review portal for <strong>' + esc(dealer) + '</strong> that summarises our conversation, the commercial terms and the Master Service Agreement.</p>',
    '<p style="margin: 20px 0;">',
    '<a href="' + esc(portalUrl) + '" style="display: inline-block; background-color: #db2777; color: #ffffff; font-weight: bold; text-decoration: none; padding: 12px 20px; border-radius: 8px;">[ Open the ' + esc(dealer) + ' Review Portal ]</a>',
    '</p>',
    '<p>The portal includes:</p>',
    '<ul>',
    '<li>' + esc(declineBullet) + '</li>',
    '<li>The fee structure: R0 retainer; R6,000 or 5% (incl. VAT), payable only per delivered vehicle</li>',
    '<li>Case studies and the compliance framework (POPIA and exclusivity)</li>',
    '<li>The Master Service Agreement, which can be reviewed and signed online</li>',
    '</ul>',
    '<p>Once the agreement is signed, your team will be directed to register on the dealership portal, where your F&amp;I referral links are issued.</p>'
  ];

  if (forwardLine) {
    htmlParts.push('<p>' + esc(forwardLine) + '</p>');
  }

  htmlParts.push(
    '<p>Kind regards,<br>',
    '<strong>' + esc(CBF_CONFIG.SENDER_SIGNATURE.name) + '</strong><br>',
    esc(CBF_CONFIG.SENDER_SIGNATURE.company) + '<br>',
    esc(CBF_CONFIG.SENDER_SIGNATURE.contact) + '</p>',
    '</div>'
  );

  return {
    subject: subject,
    plainTextBody: plainLines.join('\n'),
    htmlBody: htmlParts.join('\n')
  };
}

function escapeHtml_(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function writeTrackingCells_(sheet, rowNum, colMap, payload) {
  if (colMap['Follow-Up Status']) sheet.getRange(rowNum, colMap['Follow-Up Status']).setValue(payload.status);
  if (colMap['Portal URL']) sheet.getRange(rowNum, colMap['Portal URL']).setValue(payload.portalUrl);
  if (colMap['Gmail Draft ID']) sheet.getRange(rowNum, colMap['Gmail Draft ID']).setValue(payload.draftId);
  if (colMap['Notes Doc ID']) sheet.getRange(rowNum, colMap['Notes Doc ID']).setValue(payload.docId);
  if (colMap['Extracted JSON']) sheet.getRange(rowNum, colMap['Extracted JSON']).setValue(payload.extractedJson);
  if (colMap['Processed At']) sheet.getRange(rowNum, colMap['Processed At']).setValue(payload.processedAt);
}

/**
 * Installs a 10-minute time-driven trigger for processCompletedMeetings(),
 * removing any duplicate existing triggers for the same handler.
 */
function installTenMinuteTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processCompletedMeetings') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('processCompletedMeetings')
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log('[CBF] Installed 10-minute trigger for processCompletedMeetings.');
}

/**
 * Zero-side-effect dry-run check. Does NOT touch Drive, Sheets, or Gmail.
 * Run manually inside Apps Script Editor to verify URL building, sanitization, and Draft templates.
 */
function testDraftGeneratorDryRun() {
  var sampleFull = sanitizeExtractedFields_({
    isDealerMeeting: true,
    dpName: 'Shaun',
    dpEmail: null,
    declines: '30',
    marginK: '18',
    banks: ['WesBank', 'MFC'],
    concern: 'exclusivity'
  });

  var fullUrl = buildPortalUrl_(CBF_CONFIG.DEFAULT_PORTAL_BASE_URL, {
    dealer: 'Emerald Auto',
    contact: 'Nerissa',
    dp: sampleFull.dpName,
    date: formatMeetingDateShort_('2026-11-19'),
    declines: sampleFull.declines,
    margin: sampleFull.marginK,
    banks: sampleFull.banks,
    concern: sampleFull.concern
  });

  var fullEmail = buildFollowUpEmail_({
    dealer: 'Emerald Auto',
    contact: 'Nerissa',
    dpName: sampleFull.dpName,
    hasDpEmail: Boolean(sampleFull.dpEmail),
    declines: sampleFull.declines,
    portalUrl: fullUrl
  });

  var sampleMissing = sanitizeExtractedFields_({
    isDealerMeeting: true,
    dpName: null,
    dpEmail: 'dp@motique.co.za',
    declines: null,
    marginK: null,
    banks: null,
    concern: 'invalid_key'
  });

  var fallbackUrl = buildPortalUrl_(CBF_CONFIG.DEFAULT_PORTAL_BASE_URL, {
    dealer: 'Motique Cars',
    contact: 'Rona',
    dp: sampleMissing.dpName,
    date: '19 Nov',
    declines: sampleMissing.declines,
    margin: sampleMissing.marginK,
    banks: sampleMissing.banks,
    concern: sampleMissing.concern
  });

  var fallbackEmail = buildFollowUpEmail_({
    dealer: 'Motique Cars',
    contact: 'Rona',
    dpName: sampleMissing.dpName,
    hasDpEmail: Boolean(sampleMissing.dpEmail),
    declines: sampleMissing.declines,
    portalUrl: fallbackUrl
  });

  Logger.log('=== DRY RUN 1: FULL EXTRACTION ===');
  Logger.log('URL: ' + fullUrl);
  Logger.log('SUBJECT: ' + fullEmail.subject);
  Logger.log('TEXT:\n' + fullEmail.plainTextBody);

  Logger.log('=== DRY RUN 2: MISSING DECLINES + DP EMAIL CAPTURED (NEEDS REVIEW) ===');
  Logger.log('URL: ' + fallbackUrl);
  Logger.log('SUBJECT: ' + fallbackEmail.subject);
  Logger.log('TEXT:\n' + fallbackEmail.plainTextBody);

  return {
    full: { url: fullUrl, email: fullEmail, fields: sampleFull },
    fallback: { url: fallbackUrl, email: fallbackEmail, fields: sampleMissing }
  };
}
