/**
 * Generate validated construction schedules and Gantt charts from WBS sheets.
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
 *
 * PERT output:
 *   One activity-on-node diagram tab per WBS, grouped left-to-right by dependency level.
 */
const SCHED_TITLE_ROW = 1;
const SCHED_HEADER_ROW = 2;
const SCHED_TIMELINE_LABEL_ROW = 1;
const SCHED_TIMELINE_TENS_ROW = 2;
const SCHED_TIMELINE_DAYS_ROW = 3;
const SCHED_FIRST_DATA_ROW = 4;
const GANTT_FIRST_COLUMN = 9;
const GANTT_CELL_SIZE_PX = 20;
const PERT_NODE_ROW_SPACING = 5;
const PERT_NODE_COLUMN_SPACING = 7;
const PERT_NODE_HEIGHT = 3;
const PERT_NODE_WIDTH = 3;
const PERT_ARROW_COLOR = '#000000';
const PERT_ARROW_FONT_SIZE = 16;
const PERT_ARROW_START_PADDING = 1;
const PERT_ARROW_END_PADDING = 1;
const DEFAULT_WBS_SHEET_NAME = 'WBS';
const DEFAULT_SCHED_SHEET_NAME = 'Scheduling';
const DEFAULT_PERT_SHEET_NAME = 'PERT Diagram';
const WBS_SHEET_NAME_PATTERN = /(^|\b)WBS($|\b)/i;
const WBS_SHEET_NAME_REPLACEMENT_PATTERN = /WBS/ig;
const SCHEDULING_SHEET_ID_PROPERTY_PREFIX = 'schedulingSheetIdForWbs_';
const PERT_SHEET_ID_PROPERTY_PREFIX = 'pertSheetIdForWbs_';
const MAX_SHEET_NAME_LENGTH = 100;
const SCHEDULE_GENERATION_LOCK_TIMEOUT_MS = 25000;

function generateSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!generateScheduleForSpreadsheet_(ss)) {
    throw new Error(`Missing WBS sheet. Create a sheet with "WBS" in the tab name.`);
  }
}

function autoGenerateSchedule() {
  generateScheduleForSpreadsheet_(SpreadsheetApp.getActiveSpreadsheet());
}

function generateScheduleForSpreadsheet_(ss) {
  return runWithScheduleGenerationLock_(() => {
    const wbsSheets = getWbsSheets_(ss);

    if (wbsSheets.length === 0) return false;

    wbsSheets.forEach(wbs => generateScheduleForWbsSheet_(ss, wbs));
    return true;
  });
}

function generateScheduleForWbsSheetWithLock_(ss, wbs) {
  return runWithScheduleGenerationLock_(() => generateScheduleForWbsSheet_(ss, wbs));
}

function runWithScheduleGenerationLock_(callback) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(SCHEDULE_GENERATION_LOCK_TIMEOUT_MS)) {
    throw new Error('Schedule generation is still running. Please wait a moment, then try again.');
  }

  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function generateScheduleForWbsSheet_(ss, wbs) {
  const sched = getOrCreateSchedulingSheet_(ss, wbs);
  const pert = getOrCreatePertSheet_(ss, wbs);
  const lastRow = wbs.getLastRow();

  if (lastRow < 2) {
    clearSchedule_(sched);
    clearPertDiagram_(pert);
    return;
  }

  const rows = wbs.getRange(2, 1, lastRow - 1, 4).getValues();
  const activities = parseAndValidateWbs_(rows, wbs.getName());

  if (activities.length === 0) {
    clearSchedule_(sched);
    clearPertDiagram_(pert);
    return;
  }

  const orderedActivities = topologicalSort_(activities);
  const schedule = computeSchedule_(orderedActivities);

  renderSchedule_(sched, schedule);
  renderPertDiagram_(pert, schedule);
}

function getWbsSheets_(ss) {
  return ss.getSheets().filter(sheet => isWbsSheetName_(sheet.getName()));
}

function isWbsSheetName_(sheetName) {
  WBS_SHEET_NAME_PATTERN.lastIndex = 0;
  return WBS_SHEET_NAME_PATTERN.test(sheetName) && !isSchedulingSheetName_(sheetName) && !isPertSheetName_(sheetName);
}

function isSchedulingSheetName_(sheetName) {
  return /Scheduling/i.test(sheetName);
}

function isPertSheetName_(sheetName) {
  return /PERT/i.test(sheetName);
}

function getOrCreateSchedulingSheet_(ss, wbs) {
  const existingMappedSheet = getMappedSchedulingSheet_(ss, wbs);
  if (existingMappedSheet) return existingMappedSheet;

  const schedulingSheetName = getAvailableSchedulingSheetName_(ss, wbs);
  const existingSheet = ss.getSheetByName(schedulingSheetName);
  const sched = existingSheet || ss.insertSheet(schedulingSheetName, wbs.getIndex());
  saveSchedulingSheetMapping_(wbs, sched);
  return sched;
}

function getOrCreatePertSheet_(ss, wbs) {
  const existingMappedSheet = getMappedPertSheet_(ss, wbs);
  if (existingMappedSheet) return existingMappedSheet;

  const pertSheetName = getAvailablePertSheetName_(ss, wbs);
  const existingSheet = ss.getSheetByName(pertSheetName);
  const pert = existingSheet || ss.insertSheet(pertSheetName, wbs.getIndex() + 1);
  savePertSheetMapping_(wbs, pert);
  return pert;
}

function getMappedPertSheet_(ss, wbs) {
  const sheetId = PropertiesService.getDocumentProperties().getProperty(getPertSheetPropertyKey_(wbs));
  if (!sheetId) return null;

  const pert = ss.getSheetById(Number(sheetId));
  return pert && !isWbsSheetName_(pert.getName()) ? pert : null;
}

function savePertSheetMapping_(wbs, pert) {
  PropertiesService.getDocumentProperties().setProperty(getPertSheetPropertyKey_(wbs), String(pert.getSheetId()));
}

function getPertSheetPropertyKey_(wbs) {
  return `${PERT_SHEET_ID_PROPERTY_PREFIX}${wbs.getSheetId()}`;
}

function getAvailablePertSheetName_(ss, wbs) {
  const baseName = getPertSheetName_(wbs.getName());
  const baseSheet = ss.getSheetByName(baseName);
  if (!baseSheet || !isPertSheetMappedToOtherWbs_(baseSheet, wbs)) return baseName;

  for (let counter = 2; counter < 1000; counter++) {
    const suffix = ` ${counter}`;
    const candidate = `${baseName.slice(0, MAX_SHEET_NAME_LENGTH - suffix.length)}${suffix}`;
    const candidateSheet = ss.getSheetByName(candidate);
    if (!candidateSheet || !isPertSheetMappedToOtherWbs_(candidateSheet, wbs)) return candidate;
  }

  throw new Error(`Could not create a unique PERT Diagram tab for ${wbs.getName()}.`);
}

function isPertSheetMappedToOtherWbs_(pert, wbs) {
  const currentPropertyKey = getPertSheetPropertyKey_(wbs);
  const currentPertSheetId = String(pert.getSheetId());
  const properties = PropertiesService.getDocumentProperties().getProperties();

  return Object.keys(properties).some(key => {
    return key !== currentPropertyKey &&
      key.indexOf(PERT_SHEET_ID_PROPERTY_PREFIX) === 0 &&
      properties[key] === currentPertSheetId;
  });
}

function getPertSheetName_(wbsSheetName) {
  if (wbsSheetName === DEFAULT_WBS_SHEET_NAME) return DEFAULT_PERT_SHEET_NAME;

  WBS_SHEET_NAME_REPLACEMENT_PATTERN.lastIndex = 0;
  const pertSheetName = wbsSheetName.replace(WBS_SHEET_NAME_REPLACEMENT_PATTERN, 'PERT Diagram').trim();
  return truncateSheetName_(pertSheetName || DEFAULT_PERT_SHEET_NAME);
}

function getMappedSchedulingSheet_(ss, wbs) {
  const sheetId = PropertiesService.getDocumentProperties().getProperty(getSchedulingSheetPropertyKey_(wbs));
  if (!sheetId) return null;

  const sched = ss.getSheetById(Number(sheetId));
  return sched && !isWbsSheetName_(sched.getName()) ? sched : null;
}

function saveSchedulingSheetMapping_(wbs, sched) {
  PropertiesService.getDocumentProperties().setProperty(getSchedulingSheetPropertyKey_(wbs), String(sched.getSheetId()));
}

function getSchedulingSheetPropertyKey_(wbs) {
  return `${SCHEDULING_SHEET_ID_PROPERTY_PREFIX}${wbs.getSheetId()}`;
}

function getAvailableSchedulingSheetName_(ss, wbs) {
  const baseName = getSchedulingSheetName_(wbs.getName());
  const baseSheet = ss.getSheetByName(baseName);
  if (!baseSheet || !isSchedulingSheetMappedToOtherWbs_(baseSheet, wbs)) return baseName;

  for (let counter = 2; counter < 1000; counter++) {
    const suffix = ` ${counter}`;
    const candidate = `${baseName.slice(0, MAX_SHEET_NAME_LENGTH - suffix.length)}${suffix}`;
    const candidateSheet = ss.getSheetByName(candidate);
    if (!candidateSheet || !isSchedulingSheetMappedToOtherWbs_(candidateSheet, wbs)) return candidate;
  }

  throw new Error(`Could not create a unique Scheduling tab for ${wbs.getName()}.`);
}

function isSchedulingSheetMappedToOtherWbs_(sched, wbs) {
  const currentPropertyKey = getSchedulingSheetPropertyKey_(wbs);
  const currentScheduleSheetId = String(sched.getSheetId());
  const properties = PropertiesService.getDocumentProperties().getProperties();

  return Object.keys(properties).some(key => {
    return key !== currentPropertyKey &&
      key.indexOf(SCHEDULING_SHEET_ID_PROPERTY_PREFIX) === 0 &&
      properties[key] === currentScheduleSheetId;
  });
}

function getSchedulingSheetName_(wbsSheetName) {
  if (wbsSheetName === DEFAULT_WBS_SHEET_NAME) return DEFAULT_SCHED_SHEET_NAME;

  WBS_SHEET_NAME_REPLACEMENT_PATTERN.lastIndex = 0;
  const schedulingSheetName = wbsSheetName.replace(WBS_SHEET_NAME_REPLACEMENT_PATTERN, 'Scheduling').trim();
  return truncateSheetName_(schedulingSheetName || DEFAULT_SCHED_SHEET_NAME);
}

function truncateSheetName_(sheetName) {
  return sheetName.slice(0, MAX_SHEET_NAME_LENGTH);
}

function parseAndValidateWbs_(rows, wbsSheetName) {
  const activities = [];
  const idSet = new Set();
  const errors = [];

  rows.forEach((row, index) => {
    const sheetRow = index + 2;

    if (isBlankWbsRow_(row)) return;

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
      successors: successorsById.get(activity.id),
      slack: null,
      isCritical: false,
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
    scheduledActivity.slack = scheduledActivity.lateStart - scheduledActivity.earlyStart;
    scheduledActivity.isCritical = scheduledActivity.slack === 0;
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
  ensureSheetSize_(sched, SCHED_FIRST_DATA_ROW + output.length - 1, timeline.length + GANTT_FIRST_COLUMN - 1);
  trimExtraScheduleColumns_(sched, timeline.length + GANTT_FIRST_COLUMN - 1);

  const tableHeaderRange = sched.getRange(SCHED_HEADER_ROW, 1, 1, 8);
  tableHeaderRange.setValues([['Activity', 'Activity Description', 'Predecessor', 'Duration', 'Early Start', 'Early Finish', 'Late Start', 'Late Finish']]);
  sched.getRange(SCHED_FIRST_DATA_ROW, 1, output.length, 8).setValues(output);
  sched.getRange(SCHED_FIRST_DATA_ROW, 1, output.length, 1).setHorizontalAlignment('center');
  sched.getRange(SCHED_FIRST_DATA_ROW, 3, output.length, 6).setHorizontalAlignment('center');
  sched.autoResizeColumn(2);

  renderTimelineHeaders_(sched, timeline);
  tableHeaderRange
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const backgrounds = schedule.map(activity => timeline.map(day => {
    return day >= activity.earlyStart && day <= activity.earlyFinish ? '#4CAF50' : null;
  }));
  const ganttValues = schedule.map(activity => timeline.map(day => {
    const durationLabelDay = activity.earlyStart + Math.floor((activity.duration - 1) / 2);
    return day === durationLabelDay ? activity.duration : '';
  }));
  const ganttFontWeights = ganttValues.map(row => row.map(value => value === '' ? 'normal' : 'bold'));
  const ganttRange = sched.getRange(SCHED_FIRST_DATA_ROW, GANTT_FIRST_COLUMN, backgrounds.length, timeline.length);
  ganttRange
    .setValues(ganttValues)
    .setBackgrounds(backgrounds)
    .setFontWeights(ganttFontWeights)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  styleSchedule_(sched, output.length, timeline.length);
  resizeGanttCells_(sched, output.length, timeline.length);
  sched.setFrozenRows(SCHED_TIMELINE_DAYS_ROW);
  sched.setFrozenColumns(GANTT_FIRST_COLUMN - 1);
}

function renderPertDiagram_(pert, schedule) {
  pert.setFrozenRows(0);

  if (schedule.length === 0) {
    clearPertDiagram_(pert);
    return;
  }

  const layout = buildPertLayout_(schedule);
  const rowsNeeded = Math.max(8, 4 + layout.maxLane * PERT_NODE_ROW_SPACING + 4);
  const columnsNeeded = Math.max(10, 1 + (layout.maxLevel + 1) * PERT_NODE_COLUMN_SPACING);
  preparePertDiagramSheet_(pert, rowsNeeded, columnsNeeded);

  renderPertArrows_(pert, schedule, layout, rowsNeeded, columnsNeeded);

  const pertTitleRange = pert.getRange(1, 1, 1, columnsNeeded);
  breakApartOverlappingMergedRanges_(pertTitleRange);
  pertTitleRange
    .mergeAcross()
    .setValue('PERT DIAGRAM')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');
  const pertDescriptionRange = pert.getRange(2, 1, 1, columnsNeeded);
  breakApartOverlappingMergedRanges_(pertDescriptionRange);
  pertDescriptionRange
    .mergeAcross()
    .setValue('Each node shows ES, Duration, EF on top; Activity in the middle; and LS, Slack, LF on the bottom. Successor links use black directional arrows that point into the next activity node.')
    .setHorizontalAlignment('center');

  schedule.forEach(activity => {
    const position = layout.positions.get(activity.id);
    const row = getPertNodeRow_(position);
    const col = getPertNodeColumn_(position);
    renderPertNode_(pert, row, col, activity);
  });
  renderPertLegend_(pert, rowsNeeded, columnsNeeded);
  resizePertCells_(pert, rowsNeeded, columnsNeeded);
}

function buildPertLayout_(schedule) {
  const byId = new Map(schedule.map(activity => [activity.id, activity]));
  const levelById = new Map();

  function getLevel(activity) {
    if (levelById.has(activity.id)) return levelById.get(activity.id);

    const level = activity.predecessors.length === 0
      ? 0
      : Math.max(...activity.predecessors.map(predecessorId => getLevel(byId.get(predecessorId)))) + 1;
    levelById.set(activity.id, level);
    return level;
  }

  schedule.forEach(activity => getLevel(activity));

  const laneByLevel = new Map();
  const positions = new Map();
  let maxLane = 0;
  let maxLevel = 0;

  schedule.forEach(activity => {
    const level = levelById.get(activity.id);
    const lane = laneByLevel.get(level) || 0;
    laneByLevel.set(level, lane + 1);
    positions.set(activity.id, { level, lane });
    maxLane = Math.max(maxLane, lane);
    maxLevel = Math.max(maxLevel, level);
  });

  return { positions, maxLane, maxLevel };
}

function renderPertNode_(pert, row, col, activity) {
  const nodeRange = pert.getRange(row, col, 3, 3);
  breakApartOverlappingMergedRanges_(nodeRange);
  nodeRange
    .setValues([
      [activity.earlyStart, activity.duration, activity.earlyFinish],
      [activity.id, '', ''],
      [activity.lateStart, activity.slack, activity.lateFinish],
    ])
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setBackground('#ffffff')
    .setFontColor('#000000')
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  const activityRange = pert.getRange(row + 1, col, 1, 3);
  breakApartOverlappingMergedRanges_(activityRange);
  activityRange
    .mergeAcross()
    .setValue(activity.id)
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setBackground('#ffffff')
    .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}


function renderPertArrows_(pert, schedule, layout, rowsNeeded, columnsNeeded) {
  const activityById = new Map(schedule.map(activity => [activity.id, activity]));
  const incomingRouteIndexByTarget = new Map();
  const arrowGrid = createPertArrowGrid_(rowsNeeded, columnsNeeded);

  schedule.forEach(activity => {
    activity.predecessors.forEach((predecessorId, predecessorIndex) => {
      if (!incomingRouteIndexByTarget.has(activity.id)) {
        incomingRouteIndexByTarget.set(activity.id, new Map());
      }
      incomingRouteIndexByTarget.get(activity.id).set(predecessorId, predecessorIndex);
    });
  });

  schedule.forEach(activity => {
    const sourcePosition = layout.positions.get(activity.id);
    if (!sourcePosition) return;

    activity.successors.forEach((successorId, successorIndex) => {
      const successor = activityById.get(successorId);
      const targetPosition = layout.positions.get(successorId);
      if (!successor || !targetPosition) return;

      const incomingIndex = incomingRouteIndexByTarget.get(successorId).get(activity.id) || 0;
      const incomingCount = Math.max(1, successor.predecessors.length);
      drawPertSmartArrow_(arrowGrid, sourcePosition, targetPosition, successorIndex, incomingIndex, incomingCount);
    });
  });

  const arrowRange = pert.getRange(1, 1, rowsNeeded, columnsNeeded);
  arrowRange
    .setValues(arrowGrid)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setFontColor(PERT_ARROW_COLOR)
    .setFontSize(PERT_ARROW_FONT_SIZE)
    .setFontWeight('normal');
}

function createPertArrowGrid_(rowsNeeded, columnsNeeded) {
  return Array.from({ length: rowsNeeded }, () => Array.from({ length: columnsNeeded }, () => ''));
}

function drawPertSmartArrow_(arrowGrid, sourcePosition, targetPosition, successorIndex, incomingIndex, incomingCount) {
  const sourceRow = getPertNodeRow_(sourcePosition);
  const sourceCol = getPertNodeColumn_(sourcePosition);
  const targetRow = getPertNodeRow_(targetPosition);
  const targetCol = getPertNodeColumn_(targetPosition);
  const startPoint = getPertArrowStartPoint_(sourceRow, sourceCol);
  const endPoint = getPertArrowEndPoint_(targetRow, targetCol, incomingIndex, incomingCount);

  if (endPoint.col < startPoint.col) return;

  const horizontalClearance = endPoint.col - startPoint.col;
  const verticalDelta = endPoint.row - startPoint.row;

  if (verticalDelta === 0) {
    drawPertHorizontalArrowLine_(arrowGrid, startPoint.row, startPoint.col, endPoint.col);
    return;
  }

  const canDrawCleanDiagonal = Math.abs(verticalDelta) <= horizontalClearance && horizontalClearance <= PERT_NODE_COLUMN_SPACING;
  if (canDrawCleanDiagonal) {
    drawPertDiagonalArrow_(arrowGrid, startPoint, endPoint);
    return;
  }

  drawPertOrthogonalSmartArrow_(arrowGrid, startPoint, endPoint, successorIndex, incomingIndex);
}

function renderPertSmartArrow_(pert, sourcePosition, targetPosition, successorIndex, incomingIndex, incomingCount) {
  const sourceRow = getPertNodeRow_(sourcePosition);
  const sourceCol = getPertNodeColumn_(sourcePosition);
  const targetRow = getPertNodeRow_(targetPosition);
  const targetCol = getPertNodeColumn_(targetPosition);
  const startPoint = getPertArrowStartPoint_(sourceRow, sourceCol);
  const endPoint = getPertArrowEndPoint_(targetRow, targetCol, incomingIndex, incomingCount);

  if (endPoint.col < startPoint.col) return;

  const horizontalClearance = endPoint.col - startPoint.col;
  const verticalDelta = endPoint.row - startPoint.row;

  if (verticalDelta === 0) {
    renderPertHorizontalArrowLine_(pert, startPoint.row, startPoint.col, endPoint.col);
    renderPertArrowHead_(pert, endPoint.row, endPoint.col, '➜');
    return;
  }

  const canDrawCleanDiagonal = Math.abs(verticalDelta) <= horizontalClearance && horizontalClearance <= PERT_NODE_COLUMN_SPACING;
  if (canDrawCleanDiagonal) {
    renderPertDiagonalArrow_(pert, startPoint, endPoint);
    return;
  }

  renderPertOrthogonalSmartArrow_(pert, startPoint, endPoint, successorIndex, incomingIndex);
}

function getPertArrowStartPoint_(sourceRow, sourceCol) {
  return {
    row: sourceRow + Math.floor(PERT_NODE_HEIGHT / 2),
    col: sourceCol + PERT_NODE_WIDTH,
  };
}

function getPertArrowEndPoint_(targetRow, targetCol, incomingIndex, incomingCount) {
  return {
    row: getPertIncomingArrowRow_(targetRow, incomingIndex, incomingCount),
    col: targetCol - 1,
  };
}

function getPertIncomingArrowRow_(targetRow, incomingIndex, incomingCount) {
  if (incomingCount <= 1) return targetRow + Math.floor(PERT_NODE_HEIGHT / 2);

  const availableOffsets = Array.from({ length: PERT_NODE_HEIGHT }, (_, index) => index);
  if (incomingCount <= PERT_NODE_HEIGHT) {
    const step = (PERT_NODE_HEIGHT - 1) / (incomingCount - 1);
    return targetRow + Math.round(incomingIndex * step);
  }

  return targetRow + availableOffsets[incomingIndex % availableOffsets.length];
}

function renderPertOrthogonalSmartArrow_(pert, startPoint, endPoint, successorIndex, incomingIndex) {
  const startSpacerCol = startPoint.col + PERT_ARROW_START_PADDING;
  const endSpacerCol = endPoint.col - PERT_ARROW_END_PADDING;
  const routeWidth = Math.max(1, endSpacerCol - startSpacerCol + 1);
  const bendCol = Math.min(endSpacerCol, startSpacerCol + ((successorIndex + incomingIndex) % routeWidth));

  renderPertHorizontalConnector_(pert, startPoint.row, startPoint.col, bendCol);
  renderPertVerticalConnector_(pert, bendCol, startPoint.row, endPoint.row);
  renderPertHorizontalArrowLine_(pert, endPoint.row, bendCol, endPoint.col);
  renderPertArrowHead_(pert, endPoint.row, endPoint.col, '➜');
}

function drawPertOrthogonalSmartArrow_(arrowGrid, startPoint, endPoint, successorIndex, incomingIndex) {
  const startSpacerCol = startPoint.col + PERT_ARROW_START_PADDING;
  const endSpacerCol = endPoint.col - PERT_ARROW_END_PADDING;
  const routeWidth = Math.max(1, endSpacerCol - startSpacerCol + 1);
  const bendCol = Math.min(endSpacerCol, startSpacerCol + ((successorIndex + incomingIndex) % routeWidth));

  drawPertHorizontalConnector_(arrowGrid, startPoint.row, startPoint.col, bendCol);
  drawPertVerticalConnector_(arrowGrid, bendCol, startPoint.row, endPoint.row);
  drawPertHorizontalArrowLine_(arrowGrid, endPoint.row, bendCol, endPoint.col);
}

function renderPertDiagonalArrow_(pert, startPoint, endPoint) {
  const step = endPoint.row > startPoint.row ? 1 : -1;
  const glyph = step > 0 ? '╲' : '╱';
  const arrowGlyph = step > 0 ? '↘' : '↗';
  const diagonalSteps = Math.abs(endPoint.row - startPoint.row);
  if (diagonalSteps === 0) {
    renderPertHorizontalArrowLine_(pert, startPoint.row, startPoint.col, endPoint.col);
    renderPertArrowHead_(pert, endPoint.row, endPoint.col, '➜');
    return;
  }

  for (let stepIndex = 0; stepIndex < diagonalSteps; stepIndex++) {
    const row = startPoint.row + stepIndex * step;
    const col = startPoint.col + stepIndex;
    renderPertDiagonalConnector_(pert, row, col, glyph);
  }

  const arrowCol = Math.min(endPoint.col, startPoint.col + diagonalSteps);

  if (arrowCol < endPoint.col) {
    renderPertDiagonalConnector_(pert, endPoint.row, arrowCol, glyph);
    renderPertHorizontalArrowLine_(pert, endPoint.row, arrowCol, endPoint.col);
    renderPertArrowHead_(pert, endPoint.row, endPoint.col, '➜');
  } else {
    renderPertArrowHead_(pert, endPoint.row, arrowCol, arrowGlyph);
  }
}

function drawPertDiagonalArrow_(arrowGrid, startPoint, endPoint) {
  const step = endPoint.row > startPoint.row ? 1 : -1;
  const glyph = step > 0 ? '╲' : '╱';
  const arrowGlyph = step > 0 ? '↘' : '↗';
  const diagonalSteps = Math.abs(endPoint.row - startPoint.row);

  if (diagonalSteps === 0) {
    drawPertHorizontalArrowLine_(arrowGrid, startPoint.row, startPoint.col, endPoint.col);
    return;
  }

  for (let stepIndex = 0; stepIndex < diagonalSteps; stepIndex++) {
    const row = startPoint.row + stepIndex * step;
    const col = startPoint.col + stepIndex;
    setPertArrowGlyph_(arrowGrid, row, col, glyph);
  }

  const arrowCol = Math.min(endPoint.col, startPoint.col + diagonalSteps);

  if (arrowCol < endPoint.col) {
    setPertArrowGlyph_(arrowGrid, endPoint.row, arrowCol, glyph);
    drawPertHorizontalArrowLine_(arrowGrid, endPoint.row, arrowCol, endPoint.col);
  } else {
    setPertArrowGlyph_(arrowGrid, endPoint.row, arrowCol, arrowGlyph);
  }
}

function getPertNodeRow_(position) {
  return 4 + position.lane * PERT_NODE_ROW_SPACING;
}

function getPertNodeColumn_(position) {
  return 1 + position.level * PERT_NODE_COLUMN_SPACING;
}

function renderPertHorizontalArrowLine_(pert, row, startCol, arrowHeadCol) {
  if (arrowHeadCol > startCol) {
    renderPertHorizontalConnector_(pert, row, startCol, arrowHeadCol - 1);
  }
}

function drawPertHorizontalArrowLine_(arrowGrid, row, startCol, arrowHeadCol) {
  if (arrowHeadCol > startCol) {
    drawPertHorizontalConnector_(arrowGrid, row, startCol, arrowHeadCol - 1);
  }

  setPertArrowGlyph_(arrowGrid, row, arrowHeadCol, '➜');
}

function renderPertHorizontalConnector_(pert, row, startCol, endCol) {
  if (endCol < startCol) return;

  const connectorRange = pert.getRange(row, startCol, 1, endCol - startCol + 1);
  breakApartOverlappingMergedRanges_(connectorRange);
  connectorRange
    .clearContent()
    .setBorder(true, null, null, null, false, false, PERT_ARROW_COLOR, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function drawPertHorizontalConnector_(arrowGrid, row, startCol, endCol) {
  for (let col = startCol; col <= endCol; col++) {
    setPertArrowGlyph_(arrowGrid, row, col, '─');
  }
}

function renderPertVerticalConnector_(pert, col, startRow, endRow) {
  const topRow = Math.min(startRow, endRow);
  const rowCount = Math.abs(endRow - startRow) + 1;
  const connectorRange = pert.getRange(topRow, col, rowCount, 1);
  breakApartOverlappingMergedRanges_(connectorRange);
  connectorRange
    .clearContent()
    .setBorder(null, true, null, null, false, false, PERT_ARROW_COLOR, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function drawPertVerticalConnector_(arrowGrid, col, startRow, endRow) {
  const topRow = Math.min(startRow, endRow);
  const bottomRow = Math.max(startRow, endRow);

  for (let row = topRow; row <= bottomRow; row++) {
    setPertArrowGlyph_(arrowGrid, row, col, '│');
  }
}

function setPertArrowGlyph_(arrowGrid, row, col, glyph) {
  const rowIndex = row - 1;
  const colIndex = col - 1;

  if (rowIndex < 0 || rowIndex >= arrowGrid.length) return;
  if (colIndex < 0 || colIndex >= arrowGrid[rowIndex].length) return;

  const existingGlyph = arrowGrid[rowIndex][colIndex];
  arrowGrid[rowIndex][colIndex] = mergePertArrowGlyphs_(existingGlyph, glyph);
}

function mergePertArrowGlyphs_(existingGlyph, newGlyph) {
  if (!existingGlyph || existingGlyph === newGlyph) return newGlyph;
  if (existingGlyph === '➜' || existingGlyph === '↗' || existingGlyph === '↘') return existingGlyph;
  if (newGlyph === '➜' || newGlyph === '↗' || newGlyph === '↘') return newGlyph;
  if ((existingGlyph === '─' && newGlyph === '│') || (existingGlyph === '│' && newGlyph === '─')) return '┼';
  if (existingGlyph === '┼') return existingGlyph;
  return newGlyph;
}

function renderPertDiagonalConnector_(pert, row, col, glyph) {
  const connectorRange = pert.getRange(row, col);
  breakApartOverlappingMergedRanges_(connectorRange);
  connectorRange
    .clearContent()
    .setValue(glyph)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setFontColor(PERT_ARROW_COLOR)
    .setFontSize(PERT_ARROW_FONT_SIZE)
    .setFontWeight('normal');
}

function renderPertArrowHead_(pert, row, col, glyph) {
  const arrowHeadRange = pert.getRange(row, col);
  breakApartOverlappingMergedRanges_(arrowHeadRange);
  arrowHeadRange
    .setValue(glyph)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('right')
    .setFontColor(PERT_ARROW_COLOR)
    .setFontSize(PERT_ARROW_FONT_SIZE)
    .setFontWeight('normal');
}

function renderPertLegend_(pert, rowsNeeded, columnsNeeded) {
  const legendRow = rowsNeeded - 1;
  pert.getRange(legendRow, 1).setValue('Legend').setFontWeight('bold');
  const legendDescriptionRange = pert.getRange(legendRow, 2, 1, Math.max(1, columnsNeeded - 1));
  breakApartOverlappingMergedRanges_(legendDescriptionRange);
  legendDescriptionRange
    .mergeAcross()
    .setValue('Top: ES | Duration | EF; Middle: Activity; Bottom: LS | Slack | LF')
    .setWrap(true);
}

function resizePertCells_(pert, rowsNeeded, columnsNeeded) {
  pert.setColumnWidths(1, columnsNeeded, 80);
  pert.setRowHeights(1, rowsNeeded, 28);
}

function renderTimelineHeaders_(sched, timeline) {
  const timelineLabelRange = sched.getRange(SCHED_TIMELINE_LABEL_ROW, GANTT_FIRST_COLUMN, 1, timeline.length);
  breakApartOverlappingMergedRanges_(timelineLabelRange);
  timelineLabelRange.mergeAcross().setValue('NUMBER OF DAYS');

  renderTensHeaders_(sched, timeline);
  sched.getRange(SCHED_TIMELINE_DAYS_ROW, GANTT_FIRST_COLUMN, 1, timeline.length).setValues([timeline]);

  sched.getRange(SCHED_TIMELINE_LABEL_ROW, GANTT_FIRST_COLUMN, SCHED_TIMELINE_DAYS_ROW, timeline.length)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}


function renderTensHeaders_(sched, timeline) {
  const tensHeaderRange = sched.getRange(SCHED_TIMELINE_TENS_ROW, GANTT_FIRST_COLUMN, 1, timeline.length);
  breakApartOverlappingMergedRanges_(tensHeaderRange);
  tensHeaderRange.clearContent();

  for (let startDay = 1; startDay <= timeline.length; startDay += 10) {
    const endDay = Math.min(startDay + 9, timeline.length);
    const headerLabel = endDay;
    const headerRange = sched.getRange(SCHED_TIMELINE_TENS_ROW, GANTT_FIRST_COLUMN + startDay - 1, 1, endDay - startDay + 1);

    if (endDay > startDay) {
      breakApartOverlappingMergedRanges_(headerRange);
      headerRange.mergeAcross();
    }

    if (headerLabel) {
      headerRange.setValue(headerLabel);
    }
  }
}

function styleSchedule_(sched, activityCount, timelineLength) {
  const styledRowCount = activityCount + SCHED_FIRST_DATA_ROW - SCHED_HEADER_ROW;
  const tableRange = sched.getRange(SCHED_HEADER_ROW, 1, styledRowCount, 8);
  const timelineRange = sched.getRange(SCHED_TIMELINE_LABEL_ROW, GANTT_FIRST_COLUMN, activityCount + SCHED_FIRST_DATA_ROW - 1, timelineLength);

  tableRange.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  timelineRange.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sched.getRange(SCHED_FIRST_DATA_ROW, GANTT_FIRST_COLUMN, activityCount, timelineLength)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sched.getRange(SCHED_HEADER_ROW, 1, 1, 8).setBackground('#ffffff');
  sched.getRange(SCHED_FIRST_DATA_ROW, 1, activityCount, 8).setBackground('#ffffff');
}


/**
 * Automatically create/regenerate Scheduling tabs when the spreadsheet opens.
 *
 * This catches existing or newly added WBS tabs even before a user edits activity
 * rows, so a matching Scheduling tab appears as soon as the script detects a
 * sheet with "WBS" in its name.
 */
function onOpen(e) {
  const ss = e && e.source ? e.source : SpreadsheetApp.getActiveSpreadsheet();
  generateScheduleForSpreadsheet_(ss);
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

  generateScheduleForWbsSheetWithLock_(e.source || SpreadsheetApp.getActiveSpreadsheet(), editedSheet);
}

/**
 * Optional one-time setup for an installable change trigger.
 * Run this once if you also want schedule tabs to be created/regenerated after
 * structural spreadsheet changes such as adding/removing WBS tabs, rows, or columns.
 */
function installAutoScheduleTrigger() {
  const ss = SpreadsheetApp.getActive();
  const existingTrigger = ScriptApp.getProjectTriggers().some(trigger => {
    return trigger.getHandlerFunction() === 'autoGenerateSchedule' &&
      trigger.getEventType() === ScriptApp.EventType.ON_CHANGE;
  });

  if (existingTrigger) return;

  ScriptApp.newTrigger('autoGenerateSchedule')
    .forSpreadsheet(ss)
    .onChange()
    .create();
}

function resizeGanttCells_(sched, activityCount, timelineLength) {
  if (activityCount <= 0 || timelineLength <= 0) return;

  const totalRows = activityCount + SCHED_FIRST_DATA_ROW - 1;
  sched.setColumnWidths(GANTT_FIRST_COLUMN, timelineLength, GANTT_CELL_SIZE_PX);
  sched.setRowHeights(SCHED_TITLE_ROW, totalRows, GANTT_CELL_SIZE_PX);
}

function trimExtraScheduleColumns_(sheet, requiredColumns) {
  const extraColumns = sheet.getMaxColumns() - requiredColumns;

  if (extraColumns > 0) {
    sheet.deleteColumns(requiredColumns + 1, extraColumns);
  }
}

function trimExtraRows_(sheet, requiredRows) {
  const extraRows = sheet.getMaxRows() - requiredRows;

  if (extraRows > 0) {
    sheet.deleteRows(requiredRows + 1, extraRows);
  }
}

function ensureSheetSize_(sheet, requiredRows, requiredColumns) {
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  }
}

function preparePertDiagramSheet_(pert, requiredRows, requiredColumns) {
  ensureSheetSize_(pert, requiredRows, requiredColumns);
  clearSheetRange_(pert, requiredRows, requiredColumns);
  trimExtraRows_(pert, requiredRows);
  trimExtraScheduleColumns_(pert, requiredColumns);
}

function clearPertDiagram_(pert) {
  clearSheetRange_(pert, Math.min(pert.getMaxRows(), 50), Math.min(pert.getMaxColumns(), 26));
  trimExtraRows_(pert, 50);
  trimExtraScheduleColumns_(pert, 26);
}

function clearSchedule_(sched) {
  clearSheet_(sched);
}


function breakApartMergedRanges_(sheet) {
  const rows = sheet.getMaxRows();
  const cols = sheet.getMaxColumns();
  sheet.getRange(1, 1, rows, cols).getMergedRanges().forEach(range => range.breakApart());
}

function breakApartOverlappingMergedRanges_(range) {
  range.getMergedRanges().forEach(mergedRange => mergedRange.breakApart());
}

function clearSheet_(sheet) {
  clearSheetRange_(sheet, sheet.getMaxRows(), sheet.getMaxColumns());
}

function clearSheetRange_(sheet, rows, cols) {
  const boundedRows = Math.max(1, Math.min(rows, sheet.getMaxRows()));
  const boundedCols = Math.max(1, Math.min(cols, sheet.getMaxColumns()));
  const range = sheet.getRange(1, 1, boundedRows, boundedCols);
  breakApartOverlappingMergedRanges_(range);
  range
    .clearContent()
    .setBackground(null)
    .setFontWeight('normal')
    .setBorder(false, false, false, false, false, false);
}

function normalizeId_(value) {
  return String(value || '').trim();
}

function isBlankWbsRow_(row) {
  return row.every(value => normalizeId_(value) === '');
}

function parsePredecessors_(value) {
  const text = String(value || '').trim();
  if (!text || /^[-\u2013\u2014]$/.test(text)) return [];

  const predecessors = [];
  text.split(',').forEach(part => {
    const token = normalizeId_(part);
    const rangeMatch = token.match(/^([A-Za-z]+|\d+)\s*[-\u2013\u2014]\s*([A-Za-z]+|\d+)$/);

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

  return [joinPredecessorRange_(startToken, endToken)];
}

function expandNumericRange_(startId, endId) {
  const start = Number(startId);
  const end = Number(endId);
  if (!Number.isInteger(start) || !Number.isInteger(end) || start > end) {
    return [joinPredecessorRange_(startId, endId)];
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
  if (start > end) return [joinPredecessorRange_(startId, endId)];

  const expanded = [];

  for (let id = start; id <= end; id++) {
    expanded.push(numberToAlphabeticId_(id, startId));
  }

  return expanded;
}

function joinPredecessorRange_(startId, endId) {
  return `${startId}-${endId}`;
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
