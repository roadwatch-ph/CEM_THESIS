/**
 * Generate a validated construction schedule and Gantt chart from the WBS sheet.
 *
 * Expected WBS columns:
 *   A: Activity
 *   B: Activity Description
 *   C: Predecessor (dash/blank for none, comma-separated IDs for multiple)
 *   D: Duration
 *
 * Scheduling output:
 *   A-D: copied WBS details
 *   E: Early Start
 *   F: Early Finish
 *   G: Late Start
 *   H: Late Finish
 *   I onward: Gantt timeline
 */
const SCHED_HEADER_ROW = 1;
const SCHED_FIRST_DATA_ROW = SCHED_HEADER_ROW + 1;
const GANTT_FIRST_COLUMN = 9;
const GANTT_CELL_SIZE_PX = 20;
const DEFAULT_WBS_SHEET_NAME = 'WBS';
const DEFAULT_SCHED_SHEET_NAME = 'Scheduling';
const WBS_SHEET_NAME_PATTERN = /(^|\b)WBS($|\b)/i;
const WBS_SHEET_NAME_REPLACEMENT_PATTERN = /WBS/ig;

function generateSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wbsSheets = getWbsSheets_(ss);

  if (wbsSheets.length === 0) {
    throw new Error(`Missing WBS sheet. Create a sheet with "WBS" in the tab name.`);
  }

  wbsSheets.forEach(wbs => generateScheduleForWbsSheet_(ss, wbs));
}

function generateScheduleForWbsSheet_(ss, wbs) {
  const sched = getOrCreateSchedulingSheet_(ss, wbs);
  const lastRow = wbs.getLastRow();

  if (lastRow < 2) {
    clearSchedule_(sched);
    return;
  }

  const rows = wbs.getRange(2, 1, lastRow - 1, 4).getValues();
  const activities = parseAndValidateWbs_(rows, wbs.getName());
  const orderedActivities = topologicalSort_(activities);
  const schedule = computeSchedule_(orderedActivities);

  renderSchedule_(sched, schedule);
}

function getWbsSheets_(ss) {
  return ss.getSheets().filter(sheet => isWbsSheetName_(sheet.getName()));
}

function isWbsSheetName_(sheetName) {
  WBS_SHEET_NAME_PATTERN.lastIndex = 0;
  return WBS_SHEET_NAME_PATTERN.test(sheetName) && !isSchedulingSheetName_(sheetName);
}

function isSchedulingSheetName_(sheetName) {
  return /Scheduling/i.test(sheetName);
}

function getOrCreateSchedulingSheet_(ss, wbs) {
  const schedulingSheetName = getSchedulingSheetName_(wbs.getName());
  return ss.getSheetByName(schedulingSheetName) || ss.insertSheet(schedulingSheetName, wbs.getIndex());
}

function getSchedulingSheetName_(wbsSheetName) {
  if (wbsSheetName === DEFAULT_WBS_SHEET_NAME) return DEFAULT_SCHED_SHEET_NAME;

  WBS_SHEET_NAME_REPLACEMENT_PATTERN.lastIndex = 0;
  const schedulingSheetName = wbsSheetName.replace(WBS_SHEET_NAME_REPLACEMENT_PATTERN, 'Scheduling').trim();
  return schedulingSheetName || DEFAULT_SCHED_SHEET_NAME;
}

function parseAndValidateWbs_(rows, wbsSheetName) {
  const activities = [];
  const idSet = new Set();
  const errors = [];

  rows.forEach((row, index) => {
    const sheetRow = index + 2;
    const id = normalizeId_(row[0]);
    const name = String(row[1] || '').trim();
    const predecessors = parsePredecessors_(row[2]);
    const duration = Number(row[3]);

    if (!id) errors.push(`${wbsSheetName} row ${sheetRow}: missing Activity.`);
    if (id && idSet.has(id)) errors.push(`${wbsSheetName} row ${sheetRow}: duplicate Activity "${id}".`);
    if (!name) errors.push(`${wbsSheetName} row ${sheetRow}: missing Activity Description.`);
    if (!Number.isFinite(duration) || duration <= 0) {
      errors.push(`${wbsSheetName} row ${sheetRow}: Duration must be a positive number.`);
    }

    if (id) idSet.add(id);
    activities.push({ id, name, predecessors, duration, sourceRow: sheetRow });
  });

  activities.forEach(activity => {
    activity.predecessors.forEach(predecessor => {
      if (!idSet.has(predecessor)) {
        errors.push(`${wbsSheetName} row ${activity.sourceRow}: invalid predecessor "${predecessor}" for Activity "${activity.id}".`);
      }
      if (predecessor === activity.id) {
        errors.push(`${wbsSheetName} row ${activity.sourceRow}: activity cannot be its own predecessor.`);
      }
    });
  });

  if (errors.length > 0) {
    throw new Error(`WBS validation failed:\n${errors.join('\n')}`);
  }

  return activities;
}

function topologicalSort_(activities) {
  const byId = new Map(activities.map(activity => [activity.id, activity]));
  const visiting = new Set();
  const visited = new Set();
  const ordered = [];

  function visit(activity, path) {
    if (visited.has(activity.id)) return;

    if (visiting.has(activity.id)) {
      const cycleStart = path.indexOf(activity.id);
      const cyclePath = path.slice(cycleStart).concat(activity.id).join(' -> ');
      throw new Error(`Circular dependency detected: ${cyclePath}`);
    }

    visiting.add(activity.id);
    activity.predecessors.forEach(predecessorId => visit(byId.get(predecessorId), path.concat(activity.id)));
    visiting.delete(activity.id);
    visited.add(activity.id);
    ordered.push(activity);
  }

  activities.forEach(activity => visit(activity, []));
  return ordered;
}

function computeSchedule_(orderedActivities) {
  const scheduleById = new Map();
  const successorsById = new Map(orderedActivities.map(activity => [activity.id, []]));

  orderedActivities.forEach(activity => {
    activity.predecessors.forEach(predecessorId => {
      successorsById.get(predecessorId).push(activity.id);
    });
  });

  orderedActivities.forEach(activity => {
    const latestPredecessorEarlyFinish = activity.predecessors.reduce((latestFinish, predecessorId) => {
      return Math.max(latestFinish, scheduleById.get(predecessorId).earlyFinish);
    }, 0);

    const earlyStart = latestPredecessorEarlyFinish + 1;
    const earlyFinish = earlyStart + activity.duration - 1;

    scheduleById.set(activity.id, {
      id: activity.id,
      name: activity.name,
      predecessors: activity.predecessors,
      duration: activity.duration,
      earlyStart,
      earlyFinish,
      lateStart: null,
      lateFinish: null,
      sourceRow: activity.sourceRow,
    });
  });

  const projectFinish = Math.max(...Array.from(scheduleById.values()).map(activity => activity.earlyFinish));

  orderedActivities.slice().reverse().forEach(activity => {
    const scheduledActivity = scheduleById.get(activity.id);
    const successors = successorsById.get(activity.id);

    if (successors.length === 0) {
      scheduledActivity.lateFinish = projectFinish;
    } else {
      scheduledActivity.lateFinish = Math.min(...successors.map(successorId => scheduleById.get(successorId).lateStart - 1));
    }

    scheduledActivity.lateStart = scheduledActivity.lateFinish - scheduledActivity.duration + 1;
  });

  return Array.from(scheduleById.values()).sort((a, b) => a.sourceRow - b.sourceRow);
}

function renderSchedule_(sched, schedule) {
  clearSchedule_(sched);

  if (schedule.length === 0) return;

  const output = schedule.map(activity => [
    activity.id,
    activity.name,
    activity.predecessors.length ? activity.predecessors.join(',') : '-',
    activity.duration,
    activity.earlyStart,
    activity.earlyFinish,
    activity.lateStart,
    activity.lateFinish,
  ]);

  const maxFinish = Math.max(...schedule.map(activity => activity.lateFinish));
  const timeline = Array.from({ length: maxFinish }, (_, index) => index + 1);
  ensureSheetSize_(sched, SCHED_HEADER_ROW + output.length, timeline.length + GANTT_FIRST_COLUMN - 1);

  sched.getRange(SCHED_HEADER_ROW, 1, 1, 8).setValues([['Activity', 'Activity Description', 'Predecessor', 'Duration', 'Early Start', 'Early Finish', 'Late Start', 'Late Finish']]);
  sched.getRange(SCHED_FIRST_DATA_ROW, 1, output.length, 8).setValues(output);

  sched.getRange(SCHED_HEADER_ROW, GANTT_FIRST_COLUMN, 1, timeline.length).setValues([timeline]);

  const backgrounds = schedule.map(activity => timeline.map(day => {
    return day >= activity.earlyStart && day <= activity.earlyFinish ? '#4CAF50' : null;
  }));
  sched.getRange(SCHED_FIRST_DATA_ROW, GANTT_FIRST_COLUMN, backgrounds.length, timeline.length).setBackgrounds(backgrounds);
  resizeGanttCells_(sched, schedule.length + 1, timeline.length);
}

/**
 * Automatically regenerate the schedule after edits on the WBS sheet.
 *
 * Simple triggers run automatically for user edits, so this removes the need
 * to manually re-run generateSchedule after changing WBS activity data.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const editedSheet = e.range.getSheet();
  if (!isWbsSheetName_(editedSheet.getName())) return;

  generateScheduleForWbsSheet_(e.source || SpreadsheetApp.getActiveSpreadsheet(), editedSheet);
}

/**
 * Optional one-time setup for an installable change trigger.
 * Run this once if you also want schedule tabs to be created/regenerated after
 * structural spreadsheet changes such as adding/removing WBS tabs, rows, or columns.
 */
function installAutoScheduleTrigger() {
  const ss = SpreadsheetApp.getActive();
  const existingTrigger = ScriptApp.getProjectTriggers().some(trigger => {
    return trigger.getHandlerFunction() === 'generateSchedule' &&
      trigger.getEventType() === ScriptApp.EventType.ON_CHANGE;
  });

  if (existingTrigger) return;

  ScriptApp.newTrigger('generateSchedule')
    .forSpreadsheet(ss)
    .onChange()
    .create();
}

function resizeGanttCells_(sched, rowCount, columnCount) {
  if (rowCount <= 0 || columnCount <= 0) return;

  sched.setColumnWidths(GANTT_FIRST_COLUMN, columnCount, GANTT_CELL_SIZE_PX);
  sched.setRowHeights(SCHED_HEADER_ROW, rowCount, GANTT_CELL_SIZE_PX);
}

function ensureSheetSize_(sheet, requiredRows, requiredColumns) {
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  }
}

function clearSchedule_(sched) {
  const rows = sched.getMaxRows();
  const cols = sched.getMaxColumns();
  sched.getRange(1, 1, rows, cols).clearContent().setBackground(null);
}

function normalizeId_(value) {
  return String(value || '').trim();
}

function parsePredecessors_(value) {
  const text = String(value || '').trim();
  if (!text || text === '-') return [];

  const predecessors = [];
  text.split(',').forEach(part => {
    const token = normalizeId_(part);
    const rangeMatch = token.match(/^([A-Za-z]+|\d+)\s*-\s*([A-Za-z]+|\d+)$/);

    if (rangeMatch) {
      expandPredecessorRange_(rangeMatch[1], rangeMatch[2]).forEach(id => predecessors.push(id));
    } else {
      predecessors.push(token);
    }
  });

  return Array.from(new Set(predecessors.filter(Boolean)));
}

function expandPredecessorRange_(startId, endId) {
  const startToken = normalizeId_(startId);
  const endToken = normalizeId_(endId);

  if (/^\d+$/.test(startToken) && /^\d+$/.test(endToken)) {
    return expandNumericRange_(startToken, endToken);
  }

  if (/^[A-Za-z]+$/.test(startToken) && /^[A-Za-z]+$/.test(endToken)) {
    return expandAlphabeticRange_(startToken, endToken);
  }

  return [`${startToken}-${endToken}`];
}

function expandNumericRange_(startId, endId) {
  const start = Number(startId);
  const end = Number(endId);
  if (!Number.isInteger(start) || !Number.isInteger(end) || start > end) {
    return [`${startId}-${endId}`];
  }

  const expanded = [];

  for (let id = start; id <= end; id++) {
    expanded.push(String(id));
  }

  return expanded;
}

function expandAlphabeticRange_(startId, endId) {
  const start = alphabeticIdToNumber_(startId);
  const end = alphabeticIdToNumber_(endId);
  if (start > end) return [`${startId}-${endId}`];

  const expanded = [];

  for (let id = start; id <= end; id++) {
    expanded.push(numberToAlphabeticId_(id, startId));
  }

  return expanded;
}

function alphabeticIdToNumber_(id) {
  return id.toUpperCase().split('').reduce((total, letter) => {
    return total * 26 + letter.charCodeAt(0) - 64;
  }, 0);
}

function numberToAlphabeticId_(number, formatSource) {
  let value = number;
  let id = '';

  while (value > 0) {
    value--;
    id = String.fromCharCode(65 + (value % 26)) + id;
    value = Math.floor(value / 26);
  }

  return formatSource === formatSource.toLowerCase() ? id.toLowerCase() : id;
}
