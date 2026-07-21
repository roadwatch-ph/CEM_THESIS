/**
 * Generate validated construction schedules and Gantt charts from WBS sheets.
 *
 * Expected WBS columns:
 *   A: Activity ID (letters or numbers are supported)
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
const PERT_NODE_ROW_SPACING = 6;
const PERT_ADAPTIVE_DENSE_LEVEL_STEP = 2;
const PERT_DENSE_LEVEL_ACTIVITY_BUCKET_SIZE = 4;
const PERT_DENSE_LEVEL_DEPENDENCY_BUCKET_SIZE = 3;
const PERT_MIN_TERMINAL_ROW_SPACING = 8;
const PERT_NODE_COLUMN_SPACING = 7;
const PERT_NODE_HEIGHT = 3;
const PERT_NODE_WIDTH = 3;
const PERT_ARROW_COLOR = '#000000';
const PERT_ARROW_FONT_SIZE = 12;
const PERT_ARROW_START_PADDING = 2;
const PERT_ARROW_END_PADDING = 2;
const PERT_FIRST_NODE_ROW = 4;
const PERT_FIRST_NODE_COLUMN = 1;
const PERT_ARROW_IMAGE_ALT_TEXT = 'Generated PERT dependency arrow';
const PERT_ARROW_IMAGE_FILE_NAME = 'pert-arrow.svg';
const PERT_CELL_WIDTH_PX = 80;
const PERT_CELL_HEIGHT_PX = 28;
const PERT_ARROW_IMAGE_PADDING_PX = 4;
const PERT_ARROW_IMAGE_NODE_GAP_PX = 16;
const PERT_ARROW_BOX_CLEARANCE_PX = 8;
const PERT_MAX_ARROW_BOX_AVOIDANCE_PASSES = 12;
const PERT_MAX_ARROW_CROSSING_AVOIDANCE_PASSES = 12;
const PERT_MAX_ARROW_IMAGE_PIXELS = 2500000;
const PERT_MAX_ARROW_IMAGE_BYTES = 12000000;
const PERT_MAX_LEVELS_PER_ROW_BAND = 120;
const PERT_ROW_BAND_SPACING = 4;
const PERT_MIN_CONNECTED_NODE_ROW_DELTA = 3;
const PERT_ADAPTIVE_ROW_NUDGE = 1;
const PERT_MAX_ADAPTIVE_ROW_NUDGES_PER_NODE = 2;
const PERT_MAX_DIRECT_ARROW_RENDER_CELLS = 200000;
const PERT_MAX_IMAGE_ARROW_COUNT = 200;
const PERT_IMAGE_ARROW_MAX_NODE_COUNT = 250;
const PERT_USE_IMAGE_ARROWS = true;
const PERT_ARROW_IMAGE_STROKE_WIDTH = 2;
const PERT_ARROW_GRID_CONNECTOR_GLYPHS = new Set(['━', '┃', '┼']);
const PERT_ARROW_IMAGE_HEAD_LENGTH = 10;
const PERT_ARROW_IMAGE_HEAD_HALF_WIDTH = 6;
const PERT_ARROW_MARKER_SIZE = 12;
const PERT_WEB_ARROW_STROKE_WIDTH = 2;
const DEFAULT_WBS_SHEET_NAME = 'WBS';
const DEFAULT_SCHED_SHEET_NAME = 'Scheduling';
const DEFAULT_PERT_SHEET_NAME = 'PERT Diagram';
const PERT_START_MILESTONE_ID = 'START';
const PERT_FINISH_MILESTONE_ID = 'FINISH';
const WBS_SHEET_NAME_PATTERN = /(^|\b)WBS($|\b)/i;
const WBS_SHEET_NAME_REPLACEMENT_PATTERN = /WBS/ig;
const SCHEDULING_SHEET_ID_PROPERTY_PREFIX = 'schedulingSheetIdForWbs_';
const PERT_SHEET_ID_PROPERTY_PREFIX = 'pertSheetIdForWbs_';
const MAX_SHEET_NAME_LENGTH = 100;
const SCHEDULE_GENERATION_LOCK_TIMEOUT_MS = 25000;
const SCHEDULE_GENERATION_BUSY_MESSAGE = 'Schedule generation is already running. Please wait a moment, then try again.';
const PARALLEL_SCHEDULE_TRIGGER_HANDLER = 'runQueuedWbsScheduleGeneration_';
const PARALLEL_SCHEDULE_QUEUE_PROPERTY_PREFIX = 'parallelScheduleWbsQueue_';
const PARALLEL_SCHEDULE_SPREADSHEET_ID_PROPERTY = 'parallelScheduleSpreadsheetId';
const PARALLEL_SCHEDULE_BATCH_TIME_LIMIT_MS = 4 * 60 * 1000;
const PARALLEL_SCHEDULE_MAX_WBS_PER_RUN = 3;
const PARALLEL_SCHEDULE_TRIGGER_DELAY_MS = 1000;

function generateSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!queueParallelScheduleGenerationForSpreadsheet_(ss)) {
    throw new Error(`Missing WBS sheet. Create a sheet with "WBS" in the tab name.`);
  }
}

function generateScheduleSequential() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!generateScheduleForSpreadsheet_(ss)) {
    throw new Error(`Missing WBS sheet. Create a sheet with "WBS" in the tab name.`);
  }
}

function autoGenerateSchedule(e) {
  const ss = e && e.source ? e.source : SpreadsheetApp.getActiveSpreadsheet();

  if (shouldAutoGenerateOnlyActiveWbs_(e)) {
    const activeSheet = ss.getActiveSheet();

    if (activeSheet && isWbsSheetName_(activeSheet.getName())) {
      generateScheduleForWbsSheetWithLock_(ss, activeSheet, { skipIfBusy: true });
    }

    return;
  }

  generateScheduleForSpreadsheet_(ss, { skipIfBusy: true });
}

function shouldAutoGenerateOnlyActiveWbs_(e) {
  return e && e.changeType === 'EDIT';
}

function generateScheduleForSpreadsheet_(ss, options) {
  return runWithScheduleGenerationLock_(() => {
    const wbsSheets = getWbsSheets_(ss);

    if (wbsSheets.length === 0) return false;

    wbsSheets.forEach(wbs => generateScheduleForWbsSheet_(ss, wbs));
    return true;
  }, options);
}

function queueParallelScheduleGenerationForSpreadsheet_(ss) {
  const wbsSheets = getWbsSheets_(ss);
  if (wbsSheets.length === 0) return false;

  const properties = PropertiesService.getDocumentProperties();
  properties.setProperty(PARALLEL_SCHEDULE_SPREADSHEET_ID_PROPERTY, ss.getId());
  clearParallelScheduleQueue_();
  deleteParallelScheduleTriggers_();

  wbsSheets.forEach(wbs => {
    properties.setProperty(getParallelScheduleQueuePropertyKey_(wbs.getSheetId()), 'PENDING');
  });

  scheduleNextParallelScheduleTrigger_();
  return true;
}

function runQueuedWbsScheduleGeneration_(e) {
  deleteCurrentParallelScheduleTrigger_(e);

  const startedAt = Date.now();
  let processedCount = 0;

  while (processedCount < PARALLEL_SCHEDULE_MAX_WBS_PER_RUN && Date.now() - startedAt < PARALLEL_SCHEDULE_BATCH_TIME_LIMIT_MS) {
    const claim = claimNextQueuedWbsSchedule_({ skipIfBusy: true });
    if (!claim) break;

    try {
      const ss = SpreadsheetApp.openById(claim.spreadsheetId);
      const wbs = ss.getSheetById(claim.wbsSheetId);

      if (wbs && isWbsSheetName_(wbs.getName())) {
        generateScheduleForWbsSheet_(ss, wbs);
      }
    } finally {
      PropertiesService.getDocumentProperties().deleteProperty(claim.propertyKey);
    }

    processedCount++;
  }

  if (hasPendingParallelScheduleQueue_()) scheduleNextParallelScheduleTrigger_();
}

function claimNextQueuedWbsSchedule_(options) {
  const lock = LockService.getDocumentLock();
  const shouldSkipIfBusy = options && options.skipIfBusy;
  const hasLock = shouldSkipIfBusy
    ? lock.tryLock(1)
    : lock.tryLock(SCHEDULE_GENERATION_LOCK_TIMEOUT_MS);

  if (!hasLock) {
    if (shouldSkipIfBusy) return null;
    throw new Error(SCHEDULE_GENERATION_BUSY_MESSAGE);
  }

  try {
    const properties = PropertiesService.getDocumentProperties();
    const spreadsheetId = properties.getProperty(PARALLEL_SCHEDULE_SPREADSHEET_ID_PROPERTY);
    const allProperties = properties.getProperties();
    const propertyKey = Object.keys(allProperties).find(key => {
      return key.indexOf(PARALLEL_SCHEDULE_QUEUE_PROPERTY_PREFIX) === 0 && allProperties[key] === 'PENDING';
    });

    if (!spreadsheetId || !propertyKey) return null;

    properties.setProperty(propertyKey, 'RUNNING');
    return {
      spreadsheetId,
      propertyKey,
      wbsSheetId: Number(propertyKey.slice(PARALLEL_SCHEDULE_QUEUE_PROPERTY_PREFIX.length))
    };
  } finally {
    lock.releaseLock();
  }
}

function clearParallelScheduleQueue_() {
  const properties = PropertiesService.getDocumentProperties();
  const allProperties = properties.getProperties();

  Object.keys(allProperties).forEach(key => {
    if (key.indexOf(PARALLEL_SCHEDULE_QUEUE_PROPERTY_PREFIX) === 0) {
      properties.deleteProperty(key);
    }
  });
}

function getParallelScheduleQueuePropertyKey_(wbsSheetId) {
  return `${PARALLEL_SCHEDULE_QUEUE_PROPERTY_PREFIX}${wbsSheetId}`;
}

function hasPendingParallelScheduleQueue_() {
  const allProperties = PropertiesService.getDocumentProperties().getProperties();
  return Object.keys(allProperties).some(key => {
    return key.indexOf(PARALLEL_SCHEDULE_QUEUE_PROPERTY_PREFIX) === 0 && allProperties[key] === 'PENDING';
  });
}

function scheduleNextParallelScheduleTrigger_() {
  const hasExistingWorker = ScriptApp.getProjectTriggers().some(trigger => {
    return trigger.getHandlerFunction() === PARALLEL_SCHEDULE_TRIGGER_HANDLER;
  });
  if (hasExistingWorker) return;

  ScriptApp.newTrigger(PARALLEL_SCHEDULE_TRIGGER_HANDLER)
    .timeBased()
    .after(PARALLEL_SCHEDULE_TRIGGER_DELAY_MS)
    .create();
}

function deleteParallelScheduleTriggers_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === PARALLEL_SCHEDULE_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function deleteCurrentParallelScheduleTrigger_(e) {
  const eventUid = e && e.triggerUid ? e.triggerUid : null;
  const triggers = ScriptApp.getProjectTriggers().filter(trigger => {
    return trigger.getHandlerFunction() === PARALLEL_SCHEDULE_TRIGGER_HANDLER;
  });

  if (eventUid) {
    const currentTrigger = triggers.find(trigger => trigger.getUniqueId && trigger.getUniqueId() === eventUid);
    if (currentTrigger) {
      ScriptApp.deleteTrigger(currentTrigger);
      return;
    }
  }

  if (triggers.length > 0) {
    ScriptApp.deleteTrigger(triggers[0]);
  }
}

function generateScheduleForWbsSheetWithLock_(ss, wbs, options) {
  return runWithScheduleGenerationLock_(() => generateScheduleForWbsSheet_(ss, wbs), options);
}

function runWithScheduleGenerationLock_(callback, options) {
  const lock = LockService.getDocumentLock();
  const shouldSkipIfBusy = options && options.skipIfBusy;
  const hasLock = shouldSkipIfBusy
    ? lock.tryLock(1)
    : lock.tryLock(SCHEDULE_GENERATION_LOCK_TIMEOUT_MS);

  if (!hasLock) {
    if (shouldSkipIfBusy) return false;
    throw new Error(SCHEDULE_GENERATION_BUSY_MESSAGE);
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

  const scheduleModel = buildScheduleModel_(activities);

  renderSchedule_(sched, scheduleModel.schedule);
  renderPertDiagram_(pert, scheduleModel.schedule);
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
  const idByCanonicalId = new Map();
  const errors = [];

  rows.forEach((row, index) => {
    const sheetRow = index + 2;

    if (isBlankWbsRow_(row)) return;

    if (isWbsHeaderRow_(row)) return;

    const id = normalizeId_(row[0]);
    const canonicalId = canonicalizeActivityId_(id);
    const name = String(row[1] || '').trim();
    const predecessors = parsePredecessors_(row[2]);
    const duration = Number(row[3]);

    if (!id) errors.push(`${wbsSheetName} row ${sheetRow}: missing Activity ID.`);
    if (id && idSet.has(canonicalId)) errors.push(`${wbsSheetName} row ${sheetRow}: duplicate Activity ID "${id}".`);
    if (!name) errors.push(`${wbsSheetName} row ${sheetRow}: missing Activity Description.`);
    if (!Number.isFinite(duration) || duration <= 0) {
      errors.push(`${wbsSheetName} row ${sheetRow}: Duration must be a positive number.`);
    }

    if (id) {
      idSet.add(canonicalId);
      if (!idByCanonicalId.has(canonicalId)) idByCanonicalId.set(canonicalId, id);
    }
    activities.push({ id, name, predecessors, duration, sourceRow: sheetRow });
  });

  activities.forEach(activity => {
    activity.predecessors = activity.predecessors.map(predecessor => {
      const canonicalPredecessor = canonicalizeActivityId_(predecessor);
      return idByCanonicalId.get(canonicalPredecessor) || predecessor;
    });

    activity.predecessors.forEach(predecessor => {
      const canonicalPredecessor = canonicalizeActivityId_(predecessor);
      if (!idSet.has(canonicalPredecessor)) {
        errors.push(`${wbsSheetName} row ${activity.sourceRow}: invalid predecessor "${predecessor}" for Activity ID "${activity.id}".`);
      }
      if (canonicalPredecessor === canonicalizeActivityId_(activity.id)) {
        errors.push(`${wbsSheetName} row ${activity.sourceRow}: activity cannot be its own predecessor.`);
      }
    });
  });

  if (errors.length > 0) {
    throw new Error(`WBS validation failed:\n${errors.join('\n')}`);
  }

  return activities;
}

function buildScheduleModel_(activities) {
  const orderedActivities = topologicalSort_(activities);
  return {
    orderedActivities,
    schedule: computeSchedule_(orderedActivities),
  };
}

function topologicalSort_(activities) {
  const byId = new Map(activities.map(activity => [activity.id, activity]));
  const inDegreeById = new Map(activities.map(activity => [activity.id, activity.predecessors.length]));
  const successorsById = new Map(activities.map(activity => [activity.id, []]));
  const queue = [];
  const ordered = [];

  activities.forEach(activity => {
    activity.predecessors.forEach(predecessorId => successorsById.get(predecessorId).push(activity.id));
    if (activity.predecessors.length === 0) queue.push(activity);
  });

  queue.sort(compareActivitiesBySourceOrder_);

  while (queue.length > 0) {
    const activity = queue.shift();
    ordered.push(activity);

    successorsById.get(activity.id).forEach(successorId => {
      const nextInDegree = inDegreeById.get(successorId) - 1;
      inDegreeById.set(successorId, nextInDegree);

      if (nextInDegree === 0) {
        insertSortedActivity_(queue, byId.get(successorId), compareActivitiesBySourceOrder_);
      }
    });
  }

  if (ordered.length !== activities.length) {
    throw new Error(`Circular dependency detected: ${getDependencyCyclePath_(activities, successorsById)}`);
  }

  return ordered;
}

function compareActivitiesBySourceOrder_(a, b) {
  const sourceDelta = (a.sourceRow || 0) - (b.sourceRow || 0);
  if (sourceDelta !== 0) return sourceDelta;
  return String(a.id).localeCompare(String(b.id));
}

function insertSortedActivity_(queue, activity, compareFunction) {
  let insertIndex = queue.length;

  while (insertIndex > 0 && compareFunction(queue[insertIndex - 1], activity) > 0) {
    insertIndex--;
  }

  queue.splice(insertIndex, 0, activity);
}

function getDependencyCyclePath_(activities, successorsById) {
  const remainingIds = new Set(activities.map(activity => activity.id));
  const visited = new Set();
  const stack = new Set();
  const path = [];

  function visit(id) {
    if (stack.has(id)) {
      const cycleStart = path.indexOf(id);
      return path.slice(cycleStart).concat(id).join(' -> ');
    }

    if (visited.has(id)) return null;

    visited.add(id);
    stack.add(id);
    path.push(id);

    const successors = successorsById.get(id) || [];
    for (let index = 0; index < successors.length; index++) {
      const cyclePath = visit(successors[index]);
      if (cyclePath) return cyclePath;
    }

    stack.delete(id);
    path.pop();
    remainingIds.delete(id);
    return null;
  }

  for (let index = 0; index < activities.length; index++) {
    const cyclePath = visit(activities[index].id);
    if (cyclePath) return cyclePath;
  }

  return Array.from(remainingIds).join(' -> ') || 'unknown cycle';
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

  renderScheduleTitle_(sched);

  const tableHeaderRange = sched.getRange(SCHED_HEADER_ROW, 1, 1, 8);
  tableHeaderRange.setValues([['Activity ID', 'Activity Description', 'Predecessor', 'Duration', 'Early Start', 'Early Finish', 'Late Start', 'Late Finish']]);
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
  pert.setHiddenGridlines(true);

  if (schedule.length === 0) {
    clearPertDiagram_(pert);
    return;
  }

  const pertActivities = addPertMilestones_(schedule);
  const layout = buildPertLayout_(pertActivities);
  const rowsNeeded = Math.max(10, layout.maxNodeRow + PERT_NODE_HEIGHT + 3);
  const columnsNeeded = Math.max(12, PERT_FIRST_NODE_COLUMN + (layout.maxRenderedLevel + 1) * PERT_NODE_COLUMN_SPACING);
  preparePertDiagramSheet_(pert, rowsNeeded, columnsNeeded);
  resizePertCells_(pert, rowsNeeded, columnsNeeded);

  const pertTitleRange = pert.getRange(1, 1, 1, columnsNeeded);
  breakApartOverlappingMergedRanges_(pertTitleRange);
  pertTitleRange
    .mergeAcross()
    .setValue('PERT DIAGRAM')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setBackground('#1f4e79')
    .setFontColor('#ffffff');
  const pertDescriptionRange = pert.getRange(2, 1, 1, columnsNeeded);
  breakApartOverlappingMergedRanges_(pertDescriptionRange);
  pertDescriptionRange
    .mergeAcross()
    .setValue('Each node shows ES, Duration, EF on top; Activity ID in the middle; and LS, Slack, LF on the bottom. Arrows are rendered as spreadsheet-safe connectors with separate entry points for multiple predecessors.')
    .setHorizontalAlignment('center')
    .setWrap(true)
    .setBackground('#ddebf7');

  renderPertNodes_(pert, pertActivities, layout);

  // Place nodes before creating dependency arrows so route generation is based on
  // fixed box positions. Grid-based fallback arrows must be painted under the
  // nodes because they update spreadsheet cells across the whole diagram.
  const shouldRepaintNodes = renderPertArrows_(pert, pertActivities, layout, rowsNeeded, columnsNeeded);
  if (shouldRepaintNodes) renderPertNodes_(pert, pertActivities, layout);

  renderPertLegend_(pert, rowsNeeded, columnsNeeded);
}

function renderPertNodes_(pert, pertActivities, layout) {
  pertActivities.forEach(activity => {
    const position = layout.positions.get(activity.id);
    const row = getPertNodeRow_(position);
    const col = getPertNodeColumn_(position);
    renderPertNode_(pert, row, col, activity);
  });
}

function addPertMilestones_(schedule) {
  if (schedule.length === 0) return [];

  const projectFinish = Math.max(...schedule.map(activity => activity.earlyFinish));
  const startActivityIds = schedule
    .filter(activity => activity.predecessors.length === 0)
    .map(activity => activity.id);
  const finishActivityIds = schedule
    .filter(activity => activity.successors.length === 0)
    .map(activity => activity.id);
  const startMilestone = createPertMilestone_(PERT_START_MILESTONE_ID, [], startActivityIds, 0, 0);
  const finishMilestone = createPertMilestone_(PERT_FINISH_MILESTONE_ID, finishActivityIds, [], projectFinish + 1, schedule.length + 1);
  const activitiesWithMilestones = schedule.map(activity => {
    const hasNoPredecessors = activity.predecessors.length === 0;
    const hasNoSuccessors = activity.successors.length === 0;

    return Object.assign({}, activity, {
      predecessors: hasNoPredecessors ? [PERT_START_MILESTONE_ID] : activity.predecessors.slice(),
      successors: hasNoSuccessors ? [PERT_FINISH_MILESTONE_ID] : activity.successors.slice(),
      isPertMilestone: false,
    });
  });

  return [startMilestone].concat(activitiesWithMilestones, finishMilestone);
}

function createPertMilestone_(id, predecessors, successors, day, sourceRow) {
  return {
    id,
    name: id,
    predecessors,
    duration: 0,
    earlyStart: day,
    earlyFinish: day,
    lateStart: day,
    lateFinish: day,
    sourceRow,
    successors,
    slack: 0,
    isCritical: true,
    isPertMilestone: true,
  };
}

function buildPertLayout_(schedule) {
  const levelById = buildPertLevelMap_(schedule);
  const activitiesByLevel = new Map();
  let maxLevel = 0;

  schedule.forEach(activity => {
    const level = levelById.get(activity.id) || 0;
    if (!activitiesByLevel.has(level)) activitiesByLevel.set(level, []);
    activitiesByLevel.get(level).push(activity);
    maxLevel = Math.max(maxLevel, level);
  });

  for (let level = 0; level <= maxLevel; level++) {
    const activities = activitiesByLevel.get(level) || [];
    activities.sort(comparePertActivitiesBySourceOrder_);
    activitiesByLevel.set(level, activities);
  }

  let laneById = createPertLaneMap_(activitiesByLevel, maxLevel);
  for (let level = 1; level <= maxLevel; level++) {
    sortPertLevelByNeighborLanes_(activitiesByLevel.get(level), laneById, 'predecessors');
    laneById = createPertLaneMap_(activitiesByLevel, maxLevel);
  }

  for (let level = maxLevel - 1; level >= 0; level--) {
    sortPertLevelByNeighborLanes_(activitiesByLevel.get(level), laneById, 'successors');
    laneById = createPertLaneMap_(activitiesByLevel, maxLevel);
  }

  const positions = new Map();
  let maxLane = 0;
  for (let level = 0; level <= maxLevel; level++) {
    const activities = activitiesByLevel.get(level) || [];
    const levelTopOffsetRows = getCenteredPertLevelOffsetRows_(activities, activitiesByLevel);
    activities.forEach((activity, lane) => {
      const band = Math.floor(level / PERT_MAX_LEVELS_PER_ROW_BAND);
      const renderedLevel = level % PERT_MAX_LEVELS_PER_ROW_BAND;
      const rowBandOffset = band * getPertRowBandHeight_(activitiesByLevel);
      const levelRowSpacing = getPertAdaptiveLevelRowSpacing_(activities);
      const rowOffset = rowBandOffset + levelTopOffsetRows + lane * levelRowSpacing;
      positions.set(activity.id, { level, renderedLevel, band, lane, rowOffset });
      maxLane = Math.max(maxLane, Math.ceil(rowOffset / PERT_NODE_ROW_SPACING));
    });
  }

  alignPertFinishMilestoneRow_(schedule, positions);
  expandPertRowsForArrowClearance_(schedule, positions);
  nudgePertRowsForFanClarity_(schedule, positions);
  nudgePertRowsAwayFromDirectArrows_(schedule, positions);
  nudgePertRowsAwayFromDirectArrowCrossings_(schedule, positions);
  enforcePertNodeVerticalGaps_(positions);

  return {
    positions,
    maxLane,
    maxLevel,
    maxRenderedLevel: Math.min(maxLevel, PERT_MAX_LEVELS_PER_ROW_BAND - 1),
    maxNodeRow: getMaxPertNodeRow_(positions),
  };
}


function enforcePertNodeVerticalGaps_(positions) {
  const minimumRowGap = PERT_NODE_HEIGHT + 1;
  const positionsByRenderedColumn = new Map();

  positions.forEach(position => {
    const key = `${position.band || 0}:${position.renderedLevel}`;
    if (!positionsByRenderedColumn.has(key)) positionsByRenderedColumn.set(key, []);
    positionsByRenderedColumn.get(key).push(position);
  });

  positionsByRenderedColumn.forEach(columnPositions => {
    columnPositions
      .sort((a, b) => a.rowOffset - b.rowOffset || a.lane - b.lane)
      .reduce((previousPosition, currentPosition) => {
        if (previousPosition && currentPosition.rowOffset - previousPosition.rowOffset < minimumRowGap) {
          currentPosition.rowOffset = previousPosition.rowOffset + minimumRowGap;
        }
        return currentPosition;
      }, null);
  });
}

function alignPertFinishMilestoneRow_(schedule, positions) {
  const finishPosition = positions.get(PERT_FINISH_MILESTONE_ID);
  if (!finishPosition) return;

  const finishActivity = schedule.find(activity => activity.id === PERT_FINISH_MILESTONE_ID);
  if (!finishActivity || finishActivity.predecessors.length === 0) return;

  const predecessorRows = finishActivity.predecessors
    .map(predecessorId => positions.get(predecessorId))
    .filter(position => position)
    .map(position => position.rowOffset);

  if (predecessorRows.length === 0) return;

  const averagePredecessorRow = predecessorRows.reduce((sum, rowOffset) => sum + rowOffset, 0) / predecessorRows.length;
  finishPosition.rowOffset = Math.max(0, Math.round(averagePredecessorRow / PERT_MIN_TERMINAL_ROW_SPACING) * PERT_MIN_TERMINAL_ROW_SPACING);
}


function expandPertRowsForArrowClearance_(schedule, positions) {
  const activityById = new Map(schedule.map(activity => [activity.id, activity]));
  let changed = true;
  let passCount = 0;
  const maxPasses = Math.max(1, schedule.length * 2);

  while (changed && passCount < maxPasses) {
    changed = false;
    passCount++;

    schedule.forEach(activity => {
      const sourcePosition = positions.get(activity.id);
      if (!sourcePosition) return;

      activity.successors.forEach(successorId => {
        const successor = activityById.get(successorId);
        const targetPosition = positions.get(successorId);
        if (!successor || !targetPosition) return;

        const rowDelta = targetPosition.rowOffset - sourcePosition.rowOffset;
        if (rowDelta === 0 || Math.abs(rowDelta) >= PERT_MIN_CONNECTED_NODE_ROW_DELTA) return;

        const rowsToAdd = PERT_MIN_CONNECTED_NODE_ROW_DELTA - Math.abs(rowDelta);
        const positionToMove = rowDelta > 0 ? targetPosition : sourcePosition;
        shiftPertPositionAndLaterRows_(positions, positionToMove, rowsToAdd);
        changed = true;
      });
    });
  }
}

function shiftPertPositionAndLaterRows_(positions, anchorPosition, rowsToAdd) {
  positions.forEach(position => {
    if (position === anchorPosition || position.rowOffset >= anchorPosition.rowOffset) {
      position.rowOffset += rowsToAdd;
    }
  });
}


function nudgePertRowsForFanClarity_(schedule, positions) {
  const nudgeCountById = new Map();
  const maxPasses = Math.max(1, schedule.length);

  for (let pass = 0; pass < maxPasses; pass++) {
    const crowdedConnection = findPertCrowdedFanConnection_(schedule, positions, nudgeCountById);
    if (!crowdedConnection) return;

    nudgePertLevelRowsFrom_(positions, crowdedConnection.position, PERT_ADAPTIVE_ROW_NUDGE);
    nudgeCountById.set(crowdedConnection.id, (nudgeCountById.get(crowdedConnection.id) || 0) + 1);
  }
}

function findPertCrowdedFanConnection_(schedule, positions, nudgeCountById) {
  for (let index = 0; index < schedule.length; index++) {
    const activity = schedule[index];
    const crowdedSuccessor = findPertCrowdedNeighborPosition_(activity.successors, positions, nudgeCountById);
    if (crowdedSuccessor) return crowdedSuccessor;

    const crowdedPredecessor = findPertCrowdedNeighborPosition_(activity.predecessors, positions, nudgeCountById);
    if (crowdedPredecessor) return crowdedPredecessor;
  }

  return null;
}

function findPertCrowdedNeighborPosition_(neighborIds, positions, nudgeCountById) {
  if (!neighborIds || neighborIds.length < 2) return null;

  const neighbors = neighborIds
    .map(id => ({ id, position: positions.get(id) }))
    .filter(neighbor => neighbor.position)
    .sort((a, b) => a.position.rowOffset - b.position.rowOffset);

  for (let index = 1; index < neighbors.length; index++) {
    const previous = neighbors[index - 1];
    const current = neighbors[index];
    const rowGap = current.position.rowOffset - previous.position.rowOffset;

    if (rowGap >= PERT_MIN_CONNECTED_NODE_ROW_DELTA) continue;
    if ((nudgeCountById.get(current.id) || 0) >= PERT_MAX_ADAPTIVE_ROW_NUDGES_PER_NODE) continue;

    return current;
  }

  return null;
}

function nudgePertLevelRowsFrom_(positions, anchorPosition, rowsToAdd) {
  positions.forEach(position => {
    if (position.band !== anchorPosition.band || position.renderedLevel !== anchorPosition.renderedLevel) return;
    if (position.rowOffset >= anchorPosition.rowOffset) position.rowOffset += rowsToAdd;
  });
}

function nudgePertRowsAwayFromDirectArrows_(schedule, positions) {
  const activityById = new Map(schedule.map(activity => [activity.id, activity]));

  for (let pass = 0; pass < PERT_MAX_ARROW_BOX_AVOIDANCE_PASSES; pass++) {
    const overlap = findPertDirectArrowBoxOverlap_(schedule, activityById, positions);
    if (!overlap) return;

    if (overlap.targetId === PERT_FINISH_MILESTONE_ID) {
      nudgePertLevelRowsFrom_(positions, overlap.sourcePosition, PERT_ADAPTIVE_ROW_NUDGE);
      continue;
    }

    if (shouldLiftPertSourceForDirectArrowOverlap_(overlap)) {
      liftPertPositionWithinLevel_(positions, overlap.sourcePosition, PERT_ADAPTIVE_ROW_NUDGE);
      continue;
    }

    const positionToMove = choosePertEndpointPositionToClearOverlap_(overlap);
    nudgePertLevelRowsFrom_(positions, positionToMove, PERT_ADAPTIVE_ROW_NUDGE);
  }
}

function shouldLiftPertSourceForDirectArrowOverlap_(overlap) {
  if (overlap.sourcePosition.rowOffset <= 0) return false;

  return overlap.targetPosition.level > overlap.sourcePosition.level;
}

function choosePertEndpointPositionToClearOverlap_(overlap) {
  const sourcePosition = overlap.sourcePosition;
  const targetPosition = overlap.targetPosition;
  const blockingPosition = overlap.blockingNode.position;
  const sourceDistance = Math.abs(sourcePosition.rowOffset - blockingPosition.rowOffset);
  const targetDistance = Math.abs(targetPosition.rowOffset - blockingPosition.rowOffset);

  if (sourcePosition.rowOffset > blockingPosition.rowOffset && targetPosition.rowOffset <= blockingPosition.rowOffset) {
    return sourcePosition;
  }

  if (targetPosition.rowOffset > blockingPosition.rowOffset && sourcePosition.rowOffset <= blockingPosition.rowOffset) {
    return targetPosition;
  }

  return sourceDistance <= targetDistance ? sourcePosition : targetPosition;
}


function nudgePertRowsAwayFromDirectArrowCrossings_(schedule, positions) {
  const activityById = new Map(schedule.map(activity => [activity.id, activity]));

  for (let pass = 0; pass < PERT_MAX_ARROW_CROSSING_AVOIDANCE_PASSES; pass++) {
    const crossing = findPertDirectArrowCrossing_(schedule, activityById, positions);
    if (!crossing) return;

    const positionToLift = choosePertEndpointPositionToLiftForCrossing_(crossing);
    liftPertPositionWithinLevel_(positions, positionToLift, PERT_ADAPTIVE_ROW_NUDGE);
  }
}

function findPertDirectArrowCrossing_(schedule, activityById, positions) {
  const segments = buildPertDirectArrowSegments_(schedule, activityById, positions);

  for (let firstIndex = 0; firstIndex < segments.length; firstIndex++) {
    const first = segments[firstIndex];

    for (let secondIndex = firstIndex + 1; secondIndex < segments.length; secondIndex++) {
      const second = segments[secondIndex];
      if (doPertArrowSegmentsShareEndpoint_(first, second)) continue;
      if (!doPertLineSegmentsIntersect_(first.start, first.end, second.start, second.end)) continue;

      return { first, second };
    }
  }

  return null;
}

function buildPertDirectArrowSegments_(schedule, activityById, positions) {
  const segments = [];

  schedule.forEach(activity => {
    const sourcePosition = positions.get(activity.id);
    if (!sourcePosition) return;

    activity.successors.forEach((successorId, successorIndex) => {
      const successor = activityById.get(successorId);
      const targetPosition = positions.get(successorId);
      if (!successor || !targetPosition) return;

      const points = getPertArrowPixelConnectionPoints_(
        sourcePosition,
        targetPosition,
        successorIndex,
        activity.successors.length,
        successor.predecessors.indexOf(activity.id),
        successor.predecessors.length
      );
      applyPertArrowEndpointGap_(points.start, points.end, PERT_ARROW_IMAGE_NODE_GAP_PX);

      segments.push({
        sourceId: activity.id,
        targetId: successorId,
        sourcePosition,
        targetPosition,
        start: points.start,
        end: points.end,
      });
    });
  });

  return segments;
}

function doPertArrowSegmentsShareEndpoint_(first, second) {
  return first.sourceId === second.sourceId ||
    first.sourceId === second.targetId ||
    first.targetId === second.sourceId ||
    first.targetId === second.targetId;
}

function choosePertEndpointPositionToLiftForCrossing_(crossing) {
  if (crossing.first.targetId === PERT_FINISH_MILESTONE_ID) return crossing.first.sourcePosition;
  if (crossing.second.targetId === PERT_FINISH_MILESTONE_ID) return crossing.second.sourcePosition;

  const candidates = [
    crossing.first.sourcePosition,
    crossing.first.targetPosition,
    crossing.second.sourcePosition,
    crossing.second.targetPosition,
  ];

  return candidates.reduce((lowestPosition, position) => {
    return position.rowOffset > lowestPosition.rowOffset ? position : lowestPosition;
  }, candidates[0]);
}

function liftPertPositionWithinLevel_(positions, anchorPosition, rowsToLift) {
  anchorPosition.rowOffset = Math.max(0, anchorPosition.rowOffset - rowsToLift);
}

function findPertDirectArrowBoxOverlap_(schedule, activityById, positions) {
  for (let activityIndex = 0; activityIndex < schedule.length; activityIndex++) {
    const activity = schedule[activityIndex];
    const sourcePosition = positions.get(activity.id);
    if (!sourcePosition) continue;

    for (let successorIndex = 0; successorIndex < activity.successors.length; successorIndex++) {
      const successorId = activity.successors[successorIndex];
      const successor = activityById.get(successorId);
      const targetPosition = positions.get(successorId);
      if (!successor || !targetPosition) continue;

      const points = getPertArrowPixelConnectionPoints_(
        sourcePosition,
        targetPosition,
        successorIndex,
        activity.successors.length,
        successor.predecessors.indexOf(activity.id),
        successor.predecessors.length
      );

      applyPertArrowEndpointGap_(points.start, points.end, PERT_ARROW_IMAGE_NODE_GAP_PX);

      const blockingNode = findPertNodeBlockingDirectArrow_(positions, activity.id, successorId, points.start, points.end);
      if (blockingNode) {
        return {
          sourceId: activity.id,
          targetId: successorId,
          sourcePosition,
          targetPosition,
          blockingNode,
        };
      }
    }
  }

  return null;
}

function findPertNodeBlockingDirectArrow_(positions, sourceId, targetId, startPoint, endPoint) {
  let blockingNode = null;

  positions.forEach((position, id) => {
    if (blockingNode || id === sourceId || id === targetId) return;

    const box = expandPertPixelBox_(getPertNodePixelBox_(position), PERT_ARROW_BOX_CLEARANCE_PX);
    if (doesPertLineIntersectBox_(startPoint, endPoint, box)) {
      blockingNode = { id, position };
    }
  });

  return blockingNode;
}

function expandPertPixelBox_(box, padding) {
  return {
    left: box.left - padding,
    top: box.top - padding,
    right: box.right + padding,
    bottom: box.bottom + padding,
  };
}

function doesPertLineIntersectBox_(startPoint, endPoint, box) {
  if (isPertPointInsideBox_(startPoint, box) || isPertPointInsideBox_(endPoint, box)) return true;

  const corners = [
    { x: box.left, y: box.top },
    { x: box.right, y: box.top },
    { x: box.right, y: box.bottom },
    { x: box.left, y: box.bottom },
  ];

  for (let index = 0; index < corners.length; index++) {
    const nextIndex = (index + 1) % corners.length;
    if (doPertLineSegmentsIntersect_(startPoint, endPoint, corners[index], corners[nextIndex])) return true;
  }

  return false;
}

function isPertPointInsideBox_(point, box) {
  return point.x >= box.left && point.x <= box.right && point.y >= box.top && point.y <= box.bottom;
}

function doPertLineSegmentsIntersect_(a, b, c, d) {
  const direction1 = getPertLineDirection_(a, b, c);
  const direction2 = getPertLineDirection_(a, b, d);
  const direction3 = getPertLineDirection_(c, d, a);
  const direction4 = getPertLineDirection_(c, d, b);

  if (((direction1 > 0 && direction2 < 0) || (direction1 < 0 && direction2 > 0)) &&
      ((direction3 > 0 && direction4 < 0) || (direction3 < 0 && direction4 > 0))) {
    return true;
  }

  return direction1 === 0 && isPertPointOnLineSegment_(a, b, c) ||
    direction2 === 0 && isPertPointOnLineSegment_(a, b, d) ||
    direction3 === 0 && isPertPointOnLineSegment_(c, d, a) ||
    direction4 === 0 && isPertPointOnLineSegment_(c, d, b);
}

function getPertLineDirection_(a, b, c) {
  return (c.x - a.x) * (b.y - a.y) - (c.y - a.y) * (b.x - a.x);
}

function isPertPointOnLineSegment_(a, b, point) {
  return point.x >= Math.min(a.x, b.x) && point.x <= Math.max(a.x, b.x) &&
    point.y >= Math.min(a.y, b.y) && point.y <= Math.max(a.y, b.y);
}

function buildPertLevelMap_(schedule) {
  const byId = new Map(schedule.map(activity => [activity.id, activity]));
  const inDegreeById = new Map(schedule.map(activity => [activity.id, activity.predecessors.length]));
  const successorsById = new Map(schedule.map(activity => [activity.id, []]));
  const levelById = new Map();
  const queue = [];

  schedule.forEach(activity => {
    activity.predecessors.forEach(predecessorId => successorsById.get(predecessorId).push(activity.id));
    if (activity.predecessors.length === 0) {
      levelById.set(activity.id, 0);
      queue.push(activity);
    }
  });

  queue.sort(comparePertActivitiesBySourceOrder_);

  while (queue.length > 0) {
    const activity = queue.shift();
    const activityLevel = levelById.get(activity.id) || 0;

    successorsById.get(activity.id).forEach(successorId => {
      levelById.set(successorId, Math.max(levelById.get(successorId) || 0, activityLevel + 1));
      const nextInDegree = inDegreeById.get(successorId) - 1;
      inDegreeById.set(successorId, nextInDegree);

      if (nextInDegree === 0) {
        insertSortedActivity_(queue, byId.get(successorId), comparePertActivitiesBySourceOrder_);
      }
    });
  }

  return levelById;
}

function getPertRowBandHeight_(activitiesByLevel) {
  let tallestLevelHeight = PERT_NODE_ROW_SPACING + PERT_NODE_HEIGHT + PERT_ROW_BAND_SPACING;

  activitiesByLevel.forEach(activities => {
    const rowSpacing = getPertAdaptiveLevelRowSpacing_(activities);
    tallestLevelHeight = Math.max(tallestLevelHeight, Math.max(1, activities.length) * rowSpacing + PERT_NODE_HEIGHT + PERT_ROW_BAND_SPACING);
  });

  return tallestLevelHeight;
}

function getCenteredPertLevelOffsetRows_(activities, activitiesByLevel) {
  const targetLevelHeight = getPertRowBandHeight_(activitiesByLevel) - PERT_NODE_HEIGHT - PERT_ROW_BAND_SPACING;
  const currentLevelHeight = Math.max(1, activities.length) * getPertAdaptiveLevelRowSpacing_(activities);
  return Math.max(0, Math.floor((targetLevelHeight - currentLevelHeight) / 2));
}

function getPertAdaptiveLevelRowSpacing_(activities) {
  const activityCount = activities ? activities.length : 0;
  const maxDependencyCount = Math.max(0, ...(activities || []).map(activity => {
    return Math.max(activity.predecessors.length, activity.successors.length);
  }));
  const densitySteps = Math.floor(Math.max(0, activityCount - 1) / PERT_DENSE_LEVEL_ACTIVITY_BUCKET_SIZE);
  const dependencySteps = Math.floor(Math.max(0, maxDependencyCount - 1) / PERT_DENSE_LEVEL_DEPENDENCY_BUCKET_SIZE);

  return PERT_NODE_ROW_SPACING + (densitySteps + dependencySteps) * PERT_ADAPTIVE_DENSE_LEVEL_STEP;
}

function getMaxPertNodeRow_(positions) {
  let maxNodeRow = PERT_FIRST_NODE_ROW;
  positions.forEach(position => {
    maxNodeRow = Math.max(maxNodeRow, getPertNodeRow_(position));
  });
  return maxNodeRow;
}

function createPertLaneMap_(activitiesByLevel, maxLevel) {
  const laneById = new Map();
  for (let level = 0; level <= maxLevel; level++) {
    (activitiesByLevel.get(level) || []).forEach((activity, lane) => laneById.set(activity.id, lane));
  }
  return laneById;
}

function sortPertLevelByNeighborLanes_(activities, laneById, neighborKey) {
  if (!activities || activities.length < 2) return;

  activities.sort((a, b) => {
    const laneDelta = getAveragePertNeighborLane_(a, laneById, neighborKey) - getAveragePertNeighborLane_(b, laneById, neighborKey);
    if (laneDelta !== 0) return laneDelta;
    return comparePertActivitiesBySourceOrder_(a, b);
  });
}

function getAveragePertNeighborLane_(activity, laneById, neighborKey) {
  const neighborLanes = activity[neighborKey]
    .map(id => laneById.get(id))
    .filter(lane => lane !== undefined);

  if (neighborLanes.length === 0) return activity.sourceRow || 0;
  return neighborLanes.reduce((sum, lane) => sum + lane, 0) / neighborLanes.length;
}

function comparePertActivitiesBySourceOrder_(a, b) {
  const sourceDelta = (a.sourceRow || 0) - (b.sourceRow || 0);
  if (sourceDelta !== 0) return sourceDelta;
  return String(a.id).localeCompare(String(b.id));
}

function renderPertNode_(pert, row, col, activity) {
  const nodeRange = pert.getRange(row, col, 3, 3);
  const background = activity.isPertMilestone ? '#eaf2f8' : activity.isCritical ? '#fce4d6' : '#ffffff';
  const fontColor = activity.isPertMilestone ? '#1f4e79' : '#000000';
  const borderColor = activity.isCritical ? '#c00000' : '#000000';
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
    .setBackground(background)
    .setFontColor(fontColor)
    .setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);

  const activityRange = pert.getRange(row + 1, col, 1, 3);
  breakApartOverlappingMergedRanges_(activityRange);
  activityRange
    .mergeAcross()
    .setValue(activity.id)
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setBackground(background)
    .setFontColor(fontColor)
    .setBorder(true, true, true, true, null, null, borderColor, SpreadsheetApp.BorderStyle.SOLID);
}


function renderPertArrows_(pert, schedule, layout, rowsNeeded, columnsNeeded) {
  const arrowRoutes = buildPertArrowRoutes_(schedule, layout);
  if (arrowRoutes.length === 0) return false;

  const shouldUseImageArrows = shouldRenderPertImageArrows_(schedule, arrowRoutes);
  let fallbackArrowGrid = shouldUseImageArrows ? null : createPertArrowGrid_(rowsNeeded, columnsNeeded);
  let occupiedNodeCells = fallbackArrowGrid ? createPertOccupiedNodeCellSet_(layout.positions) : null;

  arrowRoutes.forEach(route => {
    const wasRenderedAsImage = shouldUseImageArrows && renderPertImageArrow_(
      pert,
      route.sourcePosition,
      route.targetPosition,
      route.successorIndex,
      route.successorCount,
      route.incomingIndex,
      route.incomingCount,
      layout.positions,
      route.sourceId,
      route.targetId
    );

    if (!wasRenderedAsImage) {
      if (!fallbackArrowGrid) {
        fallbackArrowGrid = createPertArrowGrid_(rowsNeeded, columnsNeeded);
        occupiedNodeCells = createPertOccupiedNodeCellSet_(layout.positions);
      }
      drawPertSmartArrow_(fallbackArrowGrid, route.sourcePosition, route.targetPosition, route.successorIndex, route.incomingIndex, route.incomingCount, occupiedNodeCells);
    }
  });

  if (fallbackArrowGrid) {
    renderPertArrowGrid_(pert, fallbackArrowGrid, rowsNeeded, columnsNeeded);
    return true;
  }

  return false;
}

function buildPertArrowRoutes_(schedule, layout) {
  const activityById = new Map(schedule.map(activity => [activity.id, activity]));
  const incomingRouteIndexByTarget = new Map();
  const arrowRoutes = [];

  schedule.forEach(activity => {
    const routeIndexByPredecessor = new Map();
    activity.predecessors.forEach((predecessorId, predecessorIndex) => {
      routeIndexByPredecessor.set(predecessorId, predecessorIndex);
    });
    incomingRouteIndexByTarget.set(activity.id, routeIndexByPredecessor);
  });

  schedule.forEach(activity => {
    const sourcePosition = layout.positions.get(activity.id);
    if (!sourcePosition) return;

    const successorCount = Math.max(1, activity.successors.length);
    activity.successors.forEach((successorId, successorIndex) => {
      const successor = activityById.get(successorId);
      const targetPosition = layout.positions.get(successorId);
      if (!successor || !targetPosition) return;

      const incomingIndexByPredecessor = incomingRouteIndexByTarget.get(successorId);
      arrowRoutes.push({
        sourceId: activity.id,
        targetId: successorId,
        sourcePosition,
        targetPosition,
        successorIndex,
        successorCount,
        incomingIndex: incomingIndexByPredecessor.get(activity.id) || 0,
        incomingCount: Math.max(1, successor.predecessors.length),
      });
    });
  });

  return arrowRoutes;
}

function shouldRenderPertImageArrows_(schedule, arrowRoutes) {
  return PERT_USE_IMAGE_ARROWS &&
    schedule.length <= PERT_IMAGE_ARROW_MAX_NODE_COUNT &&
    arrowRoutes.length <= PERT_MAX_IMAGE_ARROW_COUNT;
}


function renderPertArrowGrid_(pert, arrowGrid, rowsNeeded, columnsNeeded) {
  const displayGrid = createPertArrowDisplayGrid_(arrowGrid);

  if (rowsNeeded * columnsNeeded > PERT_MAX_DIRECT_ARROW_RENDER_CELLS) {
    renderPertArrowGridInChunks_(pert, displayGrid, rowsNeeded, columnsNeeded);
    stylePertArrowGridConnectors_(pert, arrowGrid);
    return;
  }

  const arrowRange = pert.getRange(1, 1, rowsNeeded, columnsNeeded);
  arrowRange
    .setValues(displayGrid)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setFontColor(PERT_ARROW_COLOR)
    .setFontSize(PERT_ARROW_FONT_SIZE)
    .setFontWeight('normal');
  stylePertArrowGridConnectors_(pert, arrowGrid);
}

function createPertArrowDisplayGrid_(arrowGrid) {
  return arrowGrid.map(row => row.map(glyph => PERT_ARROW_GRID_CONNECTOR_GLYPHS.has(glyph) ? '' : glyph));
}

function stylePertArrowGridConnectors_(pert, arrowGrid) {
  stylePertHorizontalArrowGridConnectors_(pert, arrowGrid);
  stylePertVerticalArrowGridConnectors_(pert, arrowGrid);
}

function stylePertHorizontalArrowGridConnectors_(pert, arrowGrid) {
  arrowGrid.forEach((row, rowIndex) => {
    let runStartCol = null;

    for (let colIndex = 0; colIndex <= row.length; colIndex++) {
      const glyph = colIndex < row.length ? row[colIndex] : '';
      const isHorizontalConnector = glyph === '━' || glyph === '┼';

      if (isHorizontalConnector && runStartCol === null) {
        runStartCol = colIndex + 1;
      }

      if ((!isHorizontalConnector || colIndex === row.length) && runStartCol !== null) {
        const endCol = colIndex;
        const connectorRange = pert.getRange(rowIndex + 1, runStartCol, 1, endCol - runStartCol + 1);
        connectorRange.setBorder(
          true,
          null,
          null,
          null,
          false,
          false,
          PERT_ARROW_COLOR,
          SpreadsheetApp.BorderStyle.SOLID
        );
        runStartCol = null;
      }
    }
  });
}

function stylePertVerticalArrowGridConnectors_(pert, arrowGrid) {
  if (arrowGrid.length === 0) return;

  const columnCount = arrowGrid[0].length;
  for (let colIndex = 0; colIndex < columnCount; colIndex++) {
    let runStartRow = null;

    for (let rowIndex = 0; rowIndex <= arrowGrid.length; rowIndex++) {
      const glyph = rowIndex < arrowGrid.length ? arrowGrid[rowIndex][colIndex] : '';
      const isVerticalConnector = glyph === '┃' || glyph === '┼';

      if (isVerticalConnector && runStartRow === null) {
        runStartRow = rowIndex + 1;
      }

      if ((!isVerticalConnector || rowIndex === arrowGrid.length) && runStartRow !== null) {
        const endRow = rowIndex;
        const connectorRange = pert.getRange(runStartRow, colIndex + 1, endRow - runStartRow + 1, 1);
        connectorRange.setBorder(
          null,
          true,
          null,
          null,
          false,
          false,
          PERT_ARROW_COLOR,
          SpreadsheetApp.BorderStyle.SOLID
        );
        runStartRow = null;
      }
    }
  }
}

function renderPertArrowGridInChunks_(pert, arrowGrid, rowsNeeded, columnsNeeded) {
  const chunkRows = Math.max(1, Math.floor(PERT_MAX_DIRECT_ARROW_RENDER_CELLS / columnsNeeded));

  for (let startRow = 1; startRow <= rowsNeeded; startRow += chunkRows) {
    const rowCount = Math.min(chunkRows, rowsNeeded - startRow + 1);
    const values = arrowGrid.slice(startRow - 1, startRow - 1 + rowCount);
    pert.getRange(startRow, 1, rowCount, columnsNeeded)
      .setValues(values)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setFontColor(PERT_ARROW_COLOR)
      .setFontSize(PERT_ARROW_FONT_SIZE)
      .setFontWeight('normal');
  }
}

function renderPertImageArrow_(pert, sourcePosition, targetPosition, successorIndex, successorCount, incomingIndex, incomingCount, positions, sourceId, targetId) {
  const connectionPoints = getPertArrowPixelConnectionPoints_(
    sourcePosition,
    targetPosition,
    successorIndex,
    successorCount,
    incomingIndex,
    incomingCount
  );
  const startPoint = connectionPoints.start;
  const endPoint = connectionPoints.end;
  applyPertArrowEndpointGap_(startPoint, endPoint, PERT_ARROW_IMAGE_NODE_GAP_PX);
  if (Math.round(startPoint.x) === Math.round(endPoint.x) && Math.round(startPoint.y) === Math.round(endPoint.y)) return false;

  const routePoints = getPertPreferredPixelRoutePoints_(startPoint, endPoint, successorIndex, incomingIndex, positions, sourceId, targetId);
  const routeBounds = getPertPixelRouteBounds_(routePoints);
  const minX = routeBounds.minX - PERT_ARROW_IMAGE_PADDING_PX;
  const minY = routeBounds.minY - PERT_ARROW_IMAGE_PADDING_PX;
  const maxX = routeBounds.maxX + PERT_ARROW_IMAGE_PADDING_PX;
  const maxY = routeBounds.maxY + PERT_ARROW_IMAGE_PADDING_PX;
  const imageWidth = Math.max(1, Math.ceil(maxX - minX));
  const imageHeight = Math.max(1, Math.ceil(maxY - minY));
  if (!canRenderPertArrowImage_(imageWidth, imageHeight)) return false;

  try {
    const localizedRoutePoints = routePoints.map(point => ({
      x: point.x - minX,
      y: point.y - minY,
    }));
    const blob = createPertArrowRouteSvgBlob_(imageWidth, imageHeight, localizedRoutePoints);
    const anchorCol = Math.max(1, Math.floor(minX / PERT_CELL_WIDTH_PX) + 1);
    const anchorRow = Math.max(1, Math.floor(minY / PERT_CELL_HEIGHT_PX) + 1);
    const xOffset = Math.max(0, Math.round(minX - (anchorCol - 1) * PERT_CELL_WIDTH_PX));
    const yOffset = Math.max(0, Math.round(minY - (anchorRow - 1) * PERT_CELL_HEIGHT_PX));

    pert.insertImage(blob, anchorCol, anchorRow, xOffset, yOffset)
      .setAltTextTitle(PERT_ARROW_IMAGE_ALT_TEXT)
      .setWidth(imageWidth)
      .setHeight(imageHeight);
    return true;
  } catch (error) {
    return false;
  }
}



function getPertPreferredPixelRoutePoints_(startPoint, endPoint, successorIndex, incomingIndex, positions, sourceId, targetId) {
  if (canUseDirectPertPixelRoute_(startPoint, endPoint, positions, sourceId, targetId)) {
    return [startPoint, endPoint];
  }

  return getPertOrthogonalPixelRoutePoints_(startPoint, endPoint, successorIndex, incomingIndex);
}

function canUseDirectPertPixelRoute_(startPoint, endPoint, positions, sourceId, targetId) {
  if (!positions) return true;

  let isClear = true;
  positions.forEach((position, id) => {
    if (!isClear || id === sourceId || id === targetId) return;

    const box = expandPertPixelBox_(getPertNodePixelBox_(position), PERT_ARROW_BOX_CLEARANCE_PX);
    if (doesPertLineIntersectBox_(startPoint, endPoint, box)) isClear = false;
  });

  return isClear;
}

function getPertOrthogonalPixelRoutePoints_(startPoint, endPoint, successorIndex, incomingIndex) {
  const horizontalDistance = endPoint.x - startPoint.x;
  if (Math.abs(endPoint.y - startPoint.y) < 1 || Math.abs(horizontalDistance) < 1) {
    return [startPoint, endPoint];
  }

  const routeOffset = Math.max(-2, Math.min(2, (successorIndex || 0) - (incomingIndex || 0))) * PERT_CELL_WIDTH_PX / 6;
  const bendX = startPoint.x + horizontalDistance / 2 + routeOffset;

  return [
    startPoint,
    { x: bendX, y: startPoint.y },
    { x: bendX, y: endPoint.y },
    endPoint,
  ];
}

function getPertPixelRouteBounds_(points) {
  return points.reduce((bounds, point) => ({
    minX: Math.min(bounds.minX, point.x),
    minY: Math.min(bounds.minY, point.y),
    maxX: Math.max(bounds.maxX, point.x),
    maxY: Math.max(bounds.maxY, point.y),
  }), {
    minX: points[0].x,
    minY: points[0].y,
    maxX: points[0].x,
    maxY: points[0].y,
  });
}

function canRenderPertArrowImage_(width, height) {
  const pixelCount = width * height;
  const estimatedPngBytes = 33 + 12 * 3 + 6 + height * (1 + width * 4);
  return pixelCount <= PERT_MAX_ARROW_IMAGE_PIXELS && estimatedPngBytes <= PERT_MAX_ARROW_IMAGE_BYTES;
}

function createPertArrowPngBlob_(width, height, startX, startY, endX, endY) {
  return createPertArrowRoutePngBlob_(width, height, [
    { x: startX, y: startY },
    { x: endX, y: endY },
  ]);
}

function createPertArrowRoutePngBlob_(width, height, points) {
  const rgba = createTransparentRgbaBuffer_(width, height);

  for (let index = 1; index < points.length; index++) {
    drawPertPngLine_(
      rgba,
      width,
      height,
      points[index - 1].x,
      points[index - 1].y,
      points[index].x,
      points[index].y,
      PERT_ARROW_IMAGE_STROKE_WIDTH
    );
  }

  const arrowStartPoint = points[Math.max(0, points.length - 2)];
  const arrowEndPoint = points[points.length - 1];
  drawPertPngArrowHead_(
    rgba,
    width,
    height,
    arrowStartPoint.x,
    arrowStartPoint.y,
    arrowEndPoint.x,
    arrowEndPoint.y,
    PERT_ARROW_IMAGE_HEAD_LENGTH,
    PERT_ARROW_IMAGE_HEAD_HALF_WIDTH
  );

  return Utilities.newBlob(createPngBytes_(width, height, rgba), 'image/png', 'pert-arrow.png');
}

function createTransparentRgbaBuffer_(width, height) {
  const pixelBytes = width * height * 4;
  if (typeof Uint8Array !== 'undefined') return new Uint8Array(pixelBytes);
  return Array.from({ length: pixelBytes }, () => 0);
}

function drawPertPngLine_(rgba, width, height, startX, startY, endX, endY, thickness) {
  const halfThickness = Math.max(0.5, thickness / 2);
  const minX = Math.max(0, Math.floor(Math.min(startX, endX) - halfThickness));
  const maxX = Math.min(width - 1, Math.ceil(Math.max(startX, endX) + halfThickness));
  const minY = Math.max(0, Math.floor(Math.min(startY, endY) - halfThickness));
  const maxY = Math.min(height - 1, Math.ceil(Math.max(startY, endY) + halfThickness));
  const radiusSquared = halfThickness * halfThickness;

  for (let y = minY; y <= maxY; y++) {
    for (let x = minX; x <= maxX; x++) {
      if (getSquaredDistanceToPertLineSegment_(x + 0.5, y + 0.5, startX, startY, endX, endY) <= radiusSquared) {
        setPertPngPixel_(rgba, width, height, x, y);
      }
    }
  }
}

function getSquaredDistanceToPertLineSegment_(pointX, pointY, startX, startY, endX, endY) {
  const dx = endX - startX;
  const dy = endY - startY;
  const lengthSquared = dx * dx + dy * dy;

  if (lengthSquared === 0) {
    const singlePointDx = pointX - startX;
    const singlePointDy = pointY - startY;
    return singlePointDx * singlePointDx + singlePointDy * singlePointDy;
  }

  const projection = Math.max(0, Math.min(1, ((pointX - startX) * dx + (pointY - startY) * dy) / lengthSquared));
  const closestX = startX + projection * dx;
  const closestY = startY + projection * dy;
  const distanceX = pointX - closestX;
  const distanceY = pointY - closestY;

  return distanceX * distanceX + distanceY * distanceY;
}

function drawPertPngArrowHead_(rgba, width, height, startX, startY, endX, endY, length, halfWidth) {
  const angle = Math.atan2(endY - startY, endX - startX);
  const baseX = endX - Math.cos(angle) * length;
  const baseY = endY - Math.sin(angle) * length;
  const normalX = -Math.sin(angle);
  const normalY = Math.cos(angle);
  const points = [
    { x: endX, y: endY },
    { x: baseX + normalX * halfWidth, y: baseY + normalY * halfWidth },
    { x: baseX - normalX * halfWidth, y: baseY - normalY * halfWidth },
  ];

  fillPertPngTriangle_(rgba, width, height, points);
}

function fillPertPngTriangle_(rgba, width, height, points) {
  const minX = Math.max(0, Math.floor(Math.min(points[0].x, points[1].x, points[2].x)));
  const maxX = Math.min(width - 1, Math.ceil(Math.max(points[0].x, points[1].x, points[2].x)));
  const minY = Math.max(0, Math.floor(Math.min(points[0].y, points[1].y, points[2].y)));
  const maxY = Math.min(height - 1, Math.ceil(Math.max(points[0].y, points[1].y, points[2].y)));

  for (let y = minY; y <= maxY; y++) {
    for (let x = minX; x <= maxX; x++) {
      if (isPointInPertTriangle_(x + 0.5, y + 0.5, points)) {
        setPertPngPixel_(rgba, width, height, x, y);
      }
    }
  }
}

function isPointInPertTriangle_(x, y, points) {
  const area = getPertTriangleSignedArea_(points[0], points[1], points[2]);
  const area1 = getPertTriangleSignedArea_({ x, y }, points[1], points[2]);
  const area2 = getPertTriangleSignedArea_(points[0], { x, y }, points[2]);
  const area3 = getPertTriangleSignedArea_(points[0], points[1], { x, y });
  const hasNegative = area1 < 0 || area2 < 0 || area3 < 0;
  const hasPositive = area1 > 0 || area2 > 0 || area3 > 0;

  return area < 0 ? !hasPositive : !hasNegative;
}

function getPertTriangleSignedArea_(a, b, c) {
  return (a.x - c.x) * (b.y - c.y) - (b.x - c.x) * (a.y - c.y);
}

function drawPertPngCircle_(rgba, width, height, centerX, centerY, radius) {
  const minX = Math.max(0, Math.floor(centerX - radius));
  const maxX = Math.min(width - 1, Math.ceil(centerX + radius));
  const minY = Math.max(0, Math.floor(centerY - radius));
  const maxY = Math.min(height - 1, Math.ceil(centerY + radius));
  const radiusSquared = radius * radius;

  for (let y = minY; y <= maxY; y++) {
    for (let x = minX; x <= maxX; x++) {
      const dx = x + 0.5 - centerX;
      const dy = y + 0.5 - centerY;
      if (dx * dx + dy * dy <= radiusSquared) {
        setPertPngPixel_(rgba, width, height, x, y);
      }
    }
  }
}

function setPertPngPixel_(rgba, width, height, x, y) {
  if (x < 0 || x >= width || y < 0 || y >= height) return;

  const offset = (y * width + x) * 4;
  rgba[offset] = 0;
  rgba[offset + 1] = 0;
  rgba[offset + 2] = 0;
  rgba[offset + 3] = 255;
}

function createPngBytes_(width, height, rgba) {
  const rawRows = [];
  for (let y = 0; y < height; y++) {
    rawRows.push(0);
    const rowOffset = y * width * 4;
    for (let x = 0; x < width * 4; x++) {
      rawRows.push(rgba[rowOffset + x]);
    }
  }

  return []
    .concat([137, 80, 78, 71, 13, 10, 26, 10])
    .concat(createPngChunk_('IHDR', uint32Bytes_(width).concat(uint32Bytes_(height), [8, 6, 0, 0, 0])))
    .concat(createPngChunk_('IDAT', createZlibStoredBlock_(rawRows)))
    .concat(createPngChunk_('IEND', []));
}

function createPngChunk_(type, data) {
  const typeBytes = type.split('').map(char => char.charCodeAt(0));
  const crc = crc32Bytes_(typeBytes.concat(data));

  return uint32Bytes_(data.length).concat(typeBytes, data, uint32Bytes_(crc));
}

function createZlibStoredBlock_(data) {
  const chunks = [0x78, 0x01];
  for (let offset = 0; offset < data.length; offset += 65535) {
    const block = data.slice(offset, offset + 65535);
    const isFinalBlock = offset + block.length >= data.length;
    const length = block.length;
    chunks.push(isFinalBlock ? 1 : 0, length & 0xff, (length >> 8) & 0xff, (~length) & 0xff, ((~length) >> 8) & 0xff);
    chunks.push.apply(chunks, block);
  }

  return chunks.concat(uint32Bytes_(adler32_(data)));
}

function uint32Bytes_(value) {
  const unsignedValue = value >>> 0;
  return [
    (unsignedValue >>> 24) & 0xff,
    (unsignedValue >>> 16) & 0xff,
    (unsignedValue >>> 8) & 0xff,
    unsignedValue & 0xff,
  ];
}

function adler32_(bytes) {
  let a = 1;
  let b = 0;

  bytes.forEach(byte => {
    a = (a + byte) % 65521;
    b = (b + a) % 65521;
  });

  return ((b << 16) | a) >>> 0;
}

function crc32Bytes_(bytes) {
  let crc = 0xffffffff;

  bytes.forEach(byte => {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) {
      crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
    }
  });

  return (crc ^ 0xffffffff) >>> 0;
}

function getPertArrowPixelStartPoint_(sourcePosition, successorIndex, successorCount) {
  const row = getPertNodeRow_(sourcePosition);
  const col = getPertNodeColumn_(sourcePosition);
  const outgoingRow = getPertOutgoingArrowRow_(row, successorIndex || 0, successorCount || 1);
  return {
    x: (col - 1 + PERT_NODE_WIDTH) * PERT_CELL_WIDTH_PX,
    y: (outgoingRow - 0.5) * PERT_CELL_HEIGHT_PX,
  };
}

function getPertOutgoingArrowRow_(sourceRow, successorIndex, successorCount) {
  return getPertDistributedNodeEdgeRow_(sourceRow, successorIndex, successorCount);
}

function getPertArrowPixelEndPoint_(targetPosition, incomingIndex, incomingCount) {
  const row = getPertNodeRow_(targetPosition);
  const col = getPertNodeColumn_(targetPosition);
  const incomingRow = getPertIncomingArrowRow_(row, incomingIndex, incomingCount);
  return {
    x: (col - 1) * PERT_CELL_WIDTH_PX,
    y: (incomingRow - 0.5) * PERT_CELL_HEIGHT_PX,
  };
}

function getPertArrowPixelConnectionPoints_(sourcePosition, targetPosition, successorIndex, successorCount, incomingIndex, incomingCount) {
  const sourceBox = getPertNodePixelBox_(sourcePosition);
  const targetBox = getPertNodePixelBox_(targetPosition);
  const sourceCenter = getPertNodePixelCenter_(sourceBox);
  const targetCenter = getPertNodePixelCenter_(targetBox);
  const dx = targetCenter.x - sourceCenter.x;
  const dy = targetCenter.y - sourceCenter.y;

  if (Math.abs(dx) >= Math.abs(dy)) {
    if (dx >= 0) {
      return {
        start: getPertPixelPointOnVerticalEdge_(sourceBox, 'right', successorIndex, successorCount),
        end: getPertPixelPointOnVerticalEdge_(targetBox, 'left', incomingIndex, incomingCount),
      };
    }

    return {
      start: getPertPixelPointOnVerticalEdge_(sourceBox, 'left', successorIndex, successorCount),
      end: getPertPixelPointOnVerticalEdge_(targetBox, 'right', incomingIndex, incomingCount),
    };
  }

  if (dy >= 0) {
    return {
      start: getPertPixelPointOnHorizontalEdge_(sourceBox, 'bottom', successorIndex, successorCount),
      end: getPertPixelPointOnHorizontalEdge_(targetBox, 'top', incomingIndex, incomingCount),
    };
  }

  return {
    start: getPertPixelPointOnHorizontalEdge_(sourceBox, 'top', successorIndex, successorCount),
    end: getPertPixelPointOnHorizontalEdge_(targetBox, 'bottom', incomingIndex, incomingCount),
  };
}

function getPertNodePixelBox_(position) {
  const row = getPertNodeRow_(position);
  const col = getPertNodeColumn_(position);
  return {
    left: (col - 1) * PERT_CELL_WIDTH_PX,
    top: (row - 1) * PERT_CELL_HEIGHT_PX,
    right: (col - 1 + PERT_NODE_WIDTH) * PERT_CELL_WIDTH_PX,
    bottom: (row - 1 + PERT_NODE_HEIGHT) * PERT_CELL_HEIGHT_PX,
  };
}

function getPertNodePixelCenter_(box) {
  return {
    x: (box.left + box.right) / 2,
    y: (box.top + box.bottom) / 2,
  };
}

function getPertPixelPointOnVerticalEdge_(box, edge, routeIndex, routeCount) {
  return {
    x: edge === 'right' ? box.right : box.left,
    y: getPertDistributedPixelCoordinate_(box.top, box.bottom, routeIndex, routeCount),
  };
}

function getPertPixelPointOnHorizontalEdge_(box, edge, routeIndex, routeCount) {
  return {
    x: getPertDistributedPixelCoordinate_(box.left, box.right, routeIndex, routeCount),
    y: edge === 'bottom' ? box.bottom : box.top,
  };
}

function getPertDistributedPixelCoordinate_(start, end, routeIndex, routeCount) {
  const count = Math.max(1, routeCount || 1);
  const index = Math.max(0, Math.min(count - 1, routeIndex || 0));
  if (count === 1) return (start + end) / 2;

  const padding = Math.max(8, Math.min(18, (end - start) / 5));
  return start + padding + (end - start - padding * 2) * index / (count - 1);
}

function applyPertArrowEndpointGap_(startPoint, endPoint, gap) {
  const dx = endPoint.x - startPoint.x;
  const dy = endPoint.y - startPoint.y;
  const distance = Math.sqrt(dx * dx + dy * dy);
  if (!distance || distance <= gap * 2) return;

  const offsetX = dx / distance * gap;
  const offsetY = dy / distance * gap;
  startPoint.x += offsetX;
  startPoint.y += offsetY;
  endPoint.x -= offsetX;
  endPoint.y -= offsetY;
}

function createPertArrowRouteSvgBlob_(width, height, points) {
  return Utilities.newBlob(
    createPertArrowRouteSvg_(width, height, points),
    'image/svg+xml',
    PERT_ARROW_IMAGE_FILE_NAME
  );
}

function createPertArrowRouteSvg_(width, height, points) {
  const safeWidth = Math.max(1, Math.ceil(width));
  const safeHeight = Math.max(1, Math.ceil(height));
  const pointList = points.map(point => `${formatPertSvgNumber_(point.x)},${formatPertSvgNumber_(point.y)}`).join(' ');
  return [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${safeWidth}" height="${safeHeight}" viewBox="0 0 ${safeWidth} ${safeHeight}">`,
    '<defs>',
    `<marker id="arrowhead" markerWidth="${PERT_ARROW_MARKER_SIZE}" markerHeight="${PERT_ARROW_MARKER_SIZE}" refX="11" refY="6" orient="auto" markerUnits="userSpaceOnUse">`,
    `<path d="M 0 0 L 12 6 L 0 12 z" fill="${PERT_ARROW_COLOR}"/>`,
    '</marker>',
    '</defs>',
    `<polyline points="${pointList}" fill="none" stroke="${PERT_ARROW_COLOR}" stroke-width="${PERT_ARROW_IMAGE_STROKE_WIDTH}" marker-end="url(#arrowhead)" stroke-linecap="round" stroke-linejoin="round"/>`,
    '</svg>',
  ].join('');
}

function createPertArrowSvg_(width, height, startX, startY, endX, endY) {
  return createPertArrowRouteSvg_(width, height, [
    { x: startX, y: startY },
    { x: endX, y: endY },
  ]);
}

function formatPertSvgNumber_(value) {
  return String(Math.round(value * 100) / 100);
}

function createPertArrowGrid_(rowsNeeded, columnsNeeded) {
  return Array.from({ length: rowsNeeded }, () => Array.from({ length: columnsNeeded }, () => ''));
}

function createPertOccupiedNodeCellSet_(positions) {
  const occupiedCells = new Set();

  positions.forEach(position => {
    const nodeRow = getPertNodeRow_(position);
    const nodeCol = getPertNodeColumn_(position);

    for (let rowOffset = 0; rowOffset < PERT_NODE_HEIGHT; rowOffset++) {
      for (let colOffset = 0; colOffset < PERT_NODE_WIDTH; colOffset++) {
        occupiedCells.add(getPertCellKey_(nodeRow + rowOffset, nodeCol + colOffset));
      }
    }
  });

  return occupiedCells;
}

function getPertCellKey_(row, col) {
  return `${row}:${col}`;
}

function drawPertSmartArrow_(arrowGrid, sourcePosition, targetPosition, successorIndex, incomingIndex, incomingCount, occupiedNodeCells) {
  const sourceRow = getPertNodeRow_(sourcePosition);
  const sourceCol = getPertNodeColumn_(sourcePosition);
  const targetRow = getPertNodeRow_(targetPosition);
  const targetCol = getPertNodeColumn_(targetPosition);
  const startPoint = getPertArrowStartPoint_(sourceRow, sourceCol);
  const endPoint = getPertArrowEndPoint_(targetRow, targetCol, incomingIndex, incomingCount);

  const verticalDelta = endPoint.row - startPoint.row;

  if (verticalDelta === 0 && !doesPertHorizontalRouteHitNode_(startPoint.row, startPoint.col, endPoint.col, occupiedNodeCells)) {
    drawPertHorizontalArrowLine_(arrowGrid, startPoint.row, startPoint.col, endPoint.col);
    return;
  }

  if (canDrawPertDiagonalRoute_(arrowGrid, startPoint, endPoint, occupiedNodeCells)) {
    drawPertDiagonalArrow_(arrowGrid, startPoint, endPoint);
    return;
  }

  drawPertOrthogonalSmartArrow_(arrowGrid, startPoint, endPoint, successorIndex || 0, incomingIndex || 0, occupiedNodeCells);
}


function drawPertWrappedArrow_(arrowGrid, startPoint, endPoint) {
  const lastCol = arrowGrid[0].length;
  drawPertHorizontalConnector_(arrowGrid, startPoint.row, startPoint.col, lastCol);
  drawPertVerticalConnector_(arrowGrid, lastCol, startPoint.row, endPoint.row);
  drawPertHorizontalArrowLine_(arrowGrid, endPoint.row, 1, endPoint.col);
}

function renderPertSmartArrow_(pert, sourcePosition, targetPosition, successorIndex, incomingIndex, incomingCount) {
  const sourceRow = getPertNodeRow_(sourcePosition);
  const sourceCol = getPertNodeColumn_(sourcePosition);
  const targetRow = getPertNodeRow_(targetPosition);
  const targetCol = getPertNodeColumn_(targetPosition);
  const startPoint = getPertArrowStartPoint_(sourceRow, sourceCol);
  const endPoint = getPertArrowEndPoint_(targetRow, targetCol, incomingIndex, incomingCount);

  if (endPoint.col < startPoint.col) return;

  const verticalDelta = endPoint.row - startPoint.row;

  if (verticalDelta === 0) {
    renderPertHorizontalArrowLine_(pert, startPoint.row, startPoint.col, endPoint.col);
    renderPertArrowHead_(pert, endPoint.row, endPoint.col, '➜');
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
  return getPertDistributedNodeEdgeRow_(targetRow, incomingIndex, incomingCount);
}

function getPertDistributedNodeEdgeRow_(nodeRow, routeIndex, routeCount) {
  if (routeCount <= 1) return nodeRow + Math.floor(PERT_NODE_HEIGHT / 2);

  const availableOffsets = Array.from({ length: PERT_NODE_HEIGHT }, (_, index) => index);
  if (routeCount <= PERT_NODE_HEIGHT) {
    const step = (PERT_NODE_HEIGHT - 1) / (routeCount - 1);
    return nodeRow + Math.round(routeIndex * step);
  }

  return nodeRow + availableOffsets[routeIndex % availableOffsets.length];
}

function renderPertOrthogonalSmartArrow_(pert, startPoint, endPoint, successorIndex, incomingIndex) {
  const bendCol = getPertOrthogonalBendColumn_(startPoint, endPoint, successorIndex, incomingIndex);

  renderPertHorizontalConnector_(pert, startPoint.row, startPoint.col, bendCol);
  renderPertVerticalConnector_(pert, bendCol, startPoint.row, endPoint.row);
  renderPertHorizontalArrowLine_(pert, endPoint.row, bendCol, endPoint.col);
  renderPertArrowHead_(pert, endPoint.row, endPoint.col, '➜');
}

function drawPertOrthogonalSmartArrow_(arrowGrid, startPoint, endPoint, successorIndex, incomingIndex, occupiedNodeCells) {
  const bendCol = occupiedNodeCells
    ? getPertNodeAvoidingBendColumn_(startPoint, endPoint)
    : getPertOrthogonalBendColumn_(startPoint, endPoint, successorIndex, incomingIndex);
  const routeRow = getPertOrthogonalRouteRow_(arrowGrid, startPoint, endPoint, bendCol, occupiedNodeCells);

  if (routeRow === null) {
    drawPertHorizontalConnector_(arrowGrid, startPoint.row, startPoint.col, bendCol);
    drawPertVerticalConnector_(arrowGrid, bendCol, startPoint.row, endPoint.row);
    drawPertHorizontalArrowLine_(arrowGrid, endPoint.row, bendCol, endPoint.col);
    return;
  }

  drawPertHorizontalConnector_(arrowGrid, startPoint.row, startPoint.col, bendCol);
  drawPertVerticalConnector_(arrowGrid, bendCol, startPoint.row, routeRow);
  drawPertHorizontalConnector_(arrowGrid, routeRow, bendCol, endPoint.col);
  drawPertVerticalConnector_(arrowGrid, endPoint.col, routeRow, endPoint.row);
  setPertArrowGlyph_(arrowGrid, endPoint.row, endPoint.col, '▶');
}

function getPertOrthogonalBendColumn_(startPoint, endPoint, successorIndex, incomingIndex) {
  const startSpacerCol = startPoint.col + PERT_ARROW_START_PADDING;
  const endSpacerCol = endPoint.col - PERT_ARROW_END_PADDING;
  const centerCol = Math.floor((startSpacerCol + endSpacerCol) / 2);
  const routeOffset = successorIndex - incomingIndex;
  const bendCol = centerCol + Math.max(-2, Math.min(2, routeOffset));

  return Math.max(startSpacerCol, Math.min(endSpacerCol, bendCol));
}

function getPertNodeAvoidingBendColumn_(startPoint, endPoint) {
  const startSpacerCol = startPoint.col + PERT_ARROW_START_PADDING;
  const endSpacerCol = endPoint.col - PERT_ARROW_END_PADDING;
  return Math.max(startSpacerCol, Math.min(endSpacerCol, startSpacerCol));
}

function getPertOrthogonalRouteRow_(arrowGrid, startPoint, endPoint, bendCol, occupiedNodeCells) {
  if (!occupiedNodeCells) return null;

  const topRow = Math.max(1, Math.min(startPoint.row, endPoint.row) - 1);
  const bottomRow = Math.min(arrowGrid.length, Math.max(startPoint.row, endPoint.row) + 1);
  const candidateRows = [topRow, bottomRow, startPoint.row, endPoint.row];

  for (let index = 0; index < candidateRows.length; index++) {
    const routeRow = candidateRows[index];
    if (isPertRouteClear_(routeRow, bendCol, endPoint.col, occupiedNodeCells)) return routeRow;
  }

  return null;
}

function isPertRouteClear_(row, startCol, endCol, occupiedNodeCells) {
  const minCol = Math.min(startCol, endCol);
  const maxCol = Math.max(startCol, endCol);

  for (let col = minCol; col <= maxCol; col++) {
    if (occupiedNodeCells.has(getPertCellKey_(row, col))) return false;
  }

  return true;
}

function doesPertHorizontalRouteHitNode_(row, startCol, endCol, occupiedNodeCells) {
  if (!occupiedNodeCells) return false;

  const minCol = Math.min(startCol, endCol);
  const maxCol = Math.max(startCol, endCol);

  for (let col = minCol; col <= maxCol; col++) {
    if (occupiedNodeCells.has(getPertCellKey_(row, col))) return true;
  }

  return false;
}


function canDrawPertDiagonalRoute_(arrowGrid, startPoint, endPoint, occupiedNodeCells) {
  if (endPoint.col <= startPoint.col || endPoint.row === startPoint.row) return false;

  const rowDelta = Math.abs(endPoint.row - startPoint.row);
  const colDelta = endPoint.col - startPoint.col;
  if (rowDelta > colDelta) return false;

  if (!occupiedNodeCells) return true;

  const step = endPoint.row > startPoint.row ? 1 : -1;
  for (let stepIndex = 0; stepIndex <= rowDelta; stepIndex++) {
    const row = startPoint.row + stepIndex * step;
    const col = startPoint.col + stepIndex;
    if (occupiedNodeCells.has(getPertCellKey_(row, col))) return false;
  }

  for (let col = startPoint.col + rowDelta; col <= endPoint.col; col++) {
    if (occupiedNodeCells.has(getPertCellKey_(endPoint.row, col))) return false;
  }

  return true;
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

function drawPertStraightOrDiagonalArrow_(arrowGrid, startPoint, endPoint) {
  const rowDelta = endPoint.row - startPoint.row;
  const colDelta = endPoint.col - startPoint.col;
  const steps = Math.max(Math.abs(rowDelta), Math.abs(colDelta), 1);

  for (let stepIndex = 0; stepIndex < steps; stepIndex++) {
    const row = Math.round(startPoint.row + rowDelta * stepIndex / steps);
    const col = Math.round(startPoint.col + colDelta * stepIndex / steps);
    const nextRow = Math.round(startPoint.row + rowDelta * (stepIndex + 1) / steps);
    const glyph = nextRow === row ? '━' : nextRow > row ? '╲' : '╱';
    setPertArrowGlyph_(arrowGrid, row, col, glyph);
  }

  const arrowGlyph = rowDelta === 0 ? '▶' : rowDelta > 0 ? '↘' : '↗';
  setPertArrowGlyph_(arrowGrid, endPoint.row, endPoint.col, arrowGlyph);
}

function getPertNodeRow_(position) {
  return PERT_FIRST_NODE_ROW + position.rowOffset;
}

function getPertNodeColumn_(position) {
  const renderedLevel = position.renderedLevel === undefined ? position.level : position.renderedLevel;
  return PERT_FIRST_NODE_COLUMN + renderedLevel * PERT_NODE_COLUMN_SPACING;
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

  setPertArrowGlyph_(arrowGrid, row, arrowHeadCol, '▶');
}

function renderPertHorizontalConnector_(pert, row, startCol, endCol) {
  if (endCol < startCol) return;

  const connectorRange = pert.getRange(row, startCol, 1, endCol - startCol + 1);
  breakApartOverlappingMergedRanges_(connectorRange);
  connectorRange
    .clearContent()
    .setBorder(true, null, null, null, false, false, PERT_ARROW_COLOR, SpreadsheetApp.BorderStyle.SOLID);
}

function drawPertHorizontalConnector_(arrowGrid, row, startCol, endCol) {
  for (let col = startCol; col <= endCol; col++) {
    setPertArrowGlyph_(arrowGrid, row, col, '━');
  }
}

function renderPertVerticalConnector_(pert, col, startRow, endRow) {
  const topRow = Math.min(startRow, endRow);
  const rowCount = Math.abs(endRow - startRow) + 1;
  const connectorRange = pert.getRange(topRow, col, rowCount, 1);
  breakApartOverlappingMergedRanges_(connectorRange);
  connectorRange
    .clearContent()
    .setBorder(null, true, null, null, false, false, PERT_ARROW_COLOR, SpreadsheetApp.BorderStyle.SOLID);
}

function drawPertVerticalConnector_(arrowGrid, col, startRow, endRow) {
  const topRow = Math.min(startRow, endRow);
  const bottomRow = Math.max(startRow, endRow);

  for (let row = topRow; row <= bottomRow; row++) {
    setPertArrowGlyph_(arrowGrid, row, col, '┃');
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
  if (existingGlyph === '▶' || existingGlyph === '➜' || existingGlyph === '↗' || existingGlyph === '↘') return existingGlyph;
  if (newGlyph === '▶' || newGlyph === '➜' || newGlyph === '↗' || newGlyph === '↘') return newGlyph;
  if ((existingGlyph === '━' && newGlyph === '┃') || (existingGlyph === '┃' && newGlyph === '━') || (existingGlyph === '─' && newGlyph === '│') || (existingGlyph === '│' && newGlyph === '─')) return '┼';
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
    .setValue('Top: ES | Duration | EF; Middle: Activity ID; Bottom: LS | Slack | LF; Blue nodes: START/FINISH milestones; Light orange/red border: critical path')
    .setWrap(true);
}

function resizePertCells_(pert, rowsNeeded, columnsNeeded) {
  pert.setColumnWidths(1, columnsNeeded, PERT_CELL_WIDTH_PX);
  pert.setRowHeights(1, rowsNeeded, PERT_CELL_HEIGHT_PX);
}


function renderScheduleTitle_(sched) {
  const titleRange = sched.getRange(SCHED_TITLE_ROW, 1, 1, GANTT_FIRST_COLUMN - 1);
  breakApartOverlappingMergedRanges_(titleRange);
  titleRange
    .mergeAcross()
    .setValue('PROJECT SCHEDULING')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#1f4e79')
    .setFontColor('#ffffff');
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
  addPertDiagramMenu_();
  generateScheduleForSpreadsheet_(ss, { skipIfBusy: true });
}

function addPertDiagramMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('PERT Tools')
    .addItem('Generate Scheduling and PERT Sheets in Parallel', 'generateSchedule')
    .addItem('Generate Scheduling and PERT Sheets Sequentially', 'generateScheduleSequential')
    .addItem('Install Auto Schedule Trigger', 'installAutoScheduleTrigger')
    .addToUi();
}

/**
 * Automatically regenerate the schedule after edits on the WBS sheet.
 *
 * Simple triggers run automatically for user edits, so this removes the need
 * to manually re-run generateSchedule after changing WBS activity data.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const editedRange = e.range;
  const editedSheet = editedRange.getSheet();
  if (!isWbsSheetName_(editedSheet.getName()) || !isWbsActivityDataEdit_(editedRange)) return;

  generateScheduleForWbsSheetWithLock_(e.source || SpreadsheetApp.getActiveSpreadsheet(), editedSheet, { skipIfBusy: true });
}

function isWbsActivityDataEdit_(range) {
  const firstEditedRow = range.getRow();
  const lastEditedRow = firstEditedRow + range.getNumRows() - 1;
  const firstEditedColumn = range.getColumn();
  const lastEditedColumn = firstEditedColumn + range.getNumColumns() - 1;

  return lastEditedRow >= 2 && firstEditedColumn <= 4 && lastEditedColumn >= 1;
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
  removeGeneratedPertArrowImages_(pert);
  ensureSheetSize_(pert, requiredRows, requiredColumns);
  clearSheetRange_(pert, requiredRows, requiredColumns);
  trimExtraRows_(pert, requiredRows);
  trimExtraScheduleColumns_(pert, requiredColumns);
}

function clearPertDiagram_(pert) {
  removeGeneratedPertArrowImages_(pert);
  clearSheetRange_(pert, Math.min(pert.getMaxRows(), 50), Math.min(pert.getMaxColumns(), 26));
  trimExtraRows_(pert, 50);
  trimExtraScheduleColumns_(pert, 26);
}

function removeGeneratedPertArrowImages_(pert) {
  pert.getImages().forEach(image => {
    if (image.getAltTextTitle && image.getAltTextTitle() === PERT_ARROW_IMAGE_ALT_TEXT) {
      image.remove();
    }
  });
}

function clearSchedule_(sched) {
  clearSheet_(sched);
  trimExtraRows_(sched, 50);
  trimExtraScheduleColumns_(sched, 26);
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

function canonicalizeActivityId_(value) {
  return normalizeId_(value).toUpperCase();
}

function isBlankWbsRow_(row) {
  return row.every(value => normalizeId_(value) === '');
}

function isWbsHeaderRow_(row) {
  const activityHeader = canonicalizeHeader_(row[0]);
  const descriptionHeader = canonicalizeHeader_(row[1]);
  const predecessorHeader = canonicalizeHeader_(row[2]);
  const durationHeader = canonicalizeHeader_(row[3]);

  return /^(ACTIVITY|ACTIVITYNO|ACTIVITYID)$/.test(activityHeader) &&
    /^(ACTIVITY|ACTIVITIES|ACTIVITYDESCRIPTION|DESCRIPTION)$/.test(descriptionHeader) &&
    /^PREDECESSORS?$/.test(predecessorHeader) &&
    /^DURATIONS?$/.test(durationHeader);
}

function canonicalizeHeader_(value) {
  return normalizeId_(value).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function parsePredecessors_(value) {
  const text = String(value || '').trim();
  if (!text || /^[-\u2013\u2014]$/.test(text)) return [];

  const predecessors = [];
  text.split(/[;,\n]+/).forEach(part => {
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

/**
 * Web app endpoint for viewing and exporting PERT diagrams outside the sheet.
 * Deploy the Apps Script project as a web app, then open the deployment URL.
 */
function doGet() {
  return HtmlService
    .createHtmlOutput(getPertDiagramWebAppHtml_())
    .setTitle('PERT Diagram Generator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getPertWebAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wbsSheets = getWbsSheets_(ss);

  return {
    spreadsheetName: ss.getName(),
    sheets: wbsSheets.map(wbs => {
      const lastRow = wbs.getLastRow();
      if (lastRow < 2) {
        return { name: wbs.getName(), activities: [], error: '' };
      }

      try {
        const rows = wbs.getRange(2, 1, lastRow - 1, 4).getValues();
        const activities = parseAndValidateWbs_(rows, wbs.getName());
        const schedule = computeSchedule_(topologicalSort_(activities));
        const pertActivities = addPertMilestones_(schedule);
        const layout = buildPertLayout_(pertActivities);
        const nodes = pertActivities.map(activity => {
          const position = layout.positions.get(activity.id);
          return Object.assign({}, activity, {
            x: (getPertNodeColumn_(position) - 1) * PERT_CELL_WIDTH_PX,
            y: (getPertNodeRow_(position) - 1) * PERT_CELL_HEIGHT_PX,
            width: PERT_NODE_WIDTH * PERT_CELL_WIDTH_PX,
            height: PERT_NODE_HEIGHT * PERT_CELL_HEIGHT_PX,
          });
        });
        const links = buildPertArrowRoutes_(pertActivities, layout).map(route => {
          const points = getPertArrowPixelConnectionPoints_(
            route.sourcePosition,
            route.targetPosition,
            route.successorIndex,
            route.successorCount,
            route.incomingIndex,
            route.incomingCount
          );
          applyPertArrowEndpointGap_(points.start, points.end, PERT_ARROW_IMAGE_NODE_GAP_PX);
          return {
            start: points.start,
            end: points.end,
            routePoints: getPertPreferredPixelRoutePoints_(
              points.start,
              points.end,
              route.successorIndex,
              route.incomingIndex,
              layout.positions,
              route.sourceId,
              route.targetId
            ),
          };
        });

        return {
          name: wbs.getName(),
          projectFinish: Math.max(...schedule.map(activity => activity.earlyFinish)),
          width: Math.max(900, (layout.maxRenderedLevel + 2) * PERT_NODE_COLUMN_SPACING * PERT_CELL_WIDTH_PX),
          height: Math.max(520, (layout.maxNodeRow + PERT_NODE_HEIGHT + 3) * PERT_CELL_HEIGHT_PX),
          activities: nodes,
          links,
          error: '',
        };
      } catch (error) {
        return { name: wbs.getName(), activities: [], error: error.message };
      }
    }),
  };
}

function getPertDiagramWebAppHtml_() {
  return `<!doctype html>
<html>
<head>
  <base target="_top">
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body { margin: 0; font-family: Arial, sans-serif; background: #f4f7fb; color: #1f2933; }
    header { background: #1f4e79; color: white; padding: 18px 24px; }
    main { padding: 18px 24px 28px; }
    .toolbar { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; margin-bottom: 14px; }
    select, button { border: 1px solid #b8c7d9; border-radius: 6px; padding: 9px 12px; background: white; font-size: 14px; }
    button { cursor: pointer; font-weight: 700; color: #1f4e79; }
    button.primary { background: #1f4e79; color: white; border-color: #1f4e79; }
    #status { color: #52606d; }
    #canvasWrap { overflow: auto; background: white; border: 1px solid #d9e2ec; border-radius: 10px; box-shadow: 0 2px 8px rgba(16, 42, 67, .08); }
    svg { display: block; min-width: 100%; }
    .node rect { stroke-width: 2; }
    .node text { text-anchor: middle; dominant-baseline: middle; font-size: 13px; }
    .node .activity { font-weight: 700; font-size: 15px; }
    .legend { margin-top: 12px; color: #52606d; font-size: 13px; }
    .error { background: #fff5f5; border: 1px solid #feb2b2; color: #9b2c2c; border-radius: 8px; padding: 12px; }
  </style>
</head>
<body>
  <header><h1>PERT Diagram Generator</h1><div id="subtitle">Loading WBS data...</div></header>
  <main>
    <div class="toolbar">
      <label>WBS Sheet <select id="sheetSelect"></select></label>
      <button class="primary" onclick="loadData()">Refresh</button>
      <button onclick="downloadSvg()">Export SVG</button>
      <button onclick="downloadPng()">Export PNG</button>
      <span id="status"></span>
    </div>
    <div id="message"></div>
    <div id="canvasWrap"><svg id="diagram" role="img" aria-label="PERT diagram"></svg></div>
    <div class="legend">Node format: ES | Duration | EF, Activity ID, LS | Slack | LF. Orange nodes are critical path activities; blue nodes are START/FINISH milestones.</div>
  </main>
<script>
let appData = null;
function loadData() {
  document.getElementById('status').textContent = 'Loading...';
  google.script.run.withSuccessHandler(data => {
    appData = data;
    document.getElementById('subtitle').textContent = data.spreadsheetName;
    const selected = document.getElementById('sheetSelect').value;
    const select = document.getElementById('sheetSelect');
    select.innerHTML = '';
    data.sheets.forEach((sheet, index) => select.add(new Option(sheet.name, index)));
    if (selected) select.value = selected;
    renderSelectedSheet();
    document.getElementById('status').textContent = 'Ready';
  }).withFailureHandler(error => {
    document.getElementById('status').textContent = 'Error';
    document.getElementById('message').innerHTML = '<div class="error">' + escapeHtml(error.message) + '</div>';
  }).getPertWebAppData();
}
function renderSelectedSheet() {
  if (!appData || appData.sheets.length === 0) return;
  const sheet = appData.sheets[Number(document.getElementById('sheetSelect').value || 0)];
  document.getElementById('message').innerHTML = sheet.error ? '<div class="error">' + escapeHtml(sheet.error) + '</div>' : '';
  renderDiagram(sheet);
}
document.getElementById('sheetSelect').addEventListener('change', renderSelectedSheet);
function renderDiagram(sheet) {
  const svg = document.getElementById('diagram');
  svg.setAttribute('viewBox', '0 0 ' + sheet.width + ' ' + sheet.height);
  svg.setAttribute('width', sheet.width);
  svg.setAttribute('height', sheet.height);
  svg.innerHTML = '<defs><marker id="arrow" markerWidth="12" markerHeight="12" refX="11" refY="6" orient="auto"><path d="M0,0 L12,6 L0,12 z" fill="#111827"/></marker></defs>';
  sheet.links.forEach(link => {
    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    const points = link.routePoints || [link.start, link.end];
    path.setAttribute('d', points.map((point, index) => (index === 0 ? 'M ' : 'L ') + point.x + ' ' + point.y).join(' '));
    path.setAttribute('fill', 'none');
    path.setAttribute('stroke', '#111827');
    path.setAttribute('stroke-width', '${PERT_WEB_ARROW_STROKE_WIDTH}');
    path.setAttribute('marker-end', 'url(#arrow)');
    svg.appendChild(path);
  });
  sheet.activities.forEach(activity => svg.appendChild(createNode(activity)));
}
function createNode(activity) {
  const group = document.createElementNS('http://www.w3.org/2000/svg', 'g');
  group.setAttribute('class', 'node');
  const fill = activity.isPertMilestone ? '#eaf2f8' : activity.isCritical ? '#fce4d6' : '#ffffff';
  const stroke = activity.isCritical ? '#c00000' : '#111827';
  const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  rect.setAttribute('x', activity.x); rect.setAttribute('y', activity.y);
  rect.setAttribute('width', activity.width); rect.setAttribute('height', activity.height);
  rect.setAttribute('rx', '8'); rect.setAttribute('fill', fill); rect.setAttribute('stroke', stroke);
  group.appendChild(rect);
  addText(group, activity.x + activity.width / 6, activity.y + activity.height / 6, activity.earlyStart);
  addText(group, activity.x + activity.width / 2, activity.y + activity.height / 6, activity.duration);
  addText(group, activity.x + activity.width * 5 / 6, activity.y + activity.height / 6, activity.earlyFinish);
  addText(group, activity.x + activity.width / 2, activity.y + activity.height / 2, activity.id, 'activity');
  addText(group, activity.x + activity.width / 6, activity.y + activity.height * 5 / 6, activity.lateStart);
  addText(group, activity.x + activity.width / 2, activity.y + activity.height * 5 / 6, activity.slack);
  addText(group, activity.x + activity.width * 5 / 6, activity.y + activity.height * 5 / 6, activity.lateFinish);
  return group;
}
function addText(group, x, y, value, className) {
  const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
  text.setAttribute('x', x); text.setAttribute('y', y);
  if (className) text.setAttribute('class', className);
  text.textContent = value;
  group.appendChild(text);
}
function downloadSvg() {
  const svg = document.getElementById('diagram').cloneNode(true);
  svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  downloadBlob(new Blob([new XMLSerializer().serializeToString(svg)], { type: 'image/svg+xml' }), fileName('svg'));
}
function downloadPng() {
  const svg = document.getElementById('diagram').cloneNode(true);
  svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  const image = new Image();
  const url = URL.createObjectURL(new Blob([new XMLSerializer().serializeToString(svg)], { type: 'image/svg+xml' }));
  image.onload = () => {
    const canvas = document.createElement('canvas');
    canvas.width = Number(svg.getAttribute('width'));
    canvas.height = Number(svg.getAttribute('height'));
    canvas.getContext('2d').fillStyle = '#ffffff';
    canvas.getContext('2d').fillRect(0, 0, canvas.width, canvas.height);
    canvas.getContext('2d').drawImage(image, 0, 0);
    URL.revokeObjectURL(url);
    canvas.toBlob(blob => downloadBlob(blob, fileName('png')));
  };
  image.src = url;
}
function downloadBlob(blob, name) {
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = name;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}
function fileName(extension) {
  const sheet = appData.sheets[Number(document.getElementById('sheetSelect').value || 0)];
  return sheet.name.replace(/[^A-Za-z0-9_-]+/g, '-') + '-PERT.' + extension;
}
function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
}
loadData();
</script>
</body>
</html>`;
}
