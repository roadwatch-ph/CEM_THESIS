/**
 * Generate a validated construction schedule and Gantt chart from the WBS sheet.
 *
 * Expected WBS columns:
 *   A: Activity No.
 *   B: Activity / Activities
 *   C: Predecessor (dash/blank for none, comma-separated IDs for multiple)
 *   D: Duration
 *
 * Scheduling output:
 *   A-D: copied WBS details
 *   E: Start Day
 *   F: Finish Day
 *   G onward: Gantt timeline
 */
function generateSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wbs = ss.getSheetByName('WBS');
  const sched = ss.getSheetByName('Scheduling');

  if (!wbs) throw new Error('Missing sheet: WBS');
  if (!sched) throw new Error('Missing sheet: Scheduling');

  const lastRow = wbs.getLastRow();
  if (lastRow < 2) {
    clearSchedule_(sched);
    return;
  }

  const rows = wbs.getRange(2, 1, lastRow - 1, 4).getValues();
  const activities = parseAndValidateWbs_(rows);
  const orderedActivities = topologicalSort_(activities);
  const schedule = computeSchedule_(orderedActivities);

  renderSchedule_(sched, schedule);
}

function parseAndValidateWbs_(rows) {
  const activities = [];
  const idSet = new Set();
  const errors = [];

  rows.forEach((row, index) => {
    const sheetRow = index + 2;
    const id = normalizeId_(row[0]);
    const name = String(row[1] || '').trim();
    const predecessors = parsePredecessors_(row[2]);
    const duration = Number(row[3]);

    if (!id) errors.push(`Row ${sheetRow}: missing Activity No.`);
    if (id && idSet.has(id)) errors.push(`Row ${sheetRow}: duplicate Activity No. "${id}".`);
    if (!name) errors.push(`Row ${sheetRow}: missing Activity name.`);
    if (!Number.isFinite(duration) || duration <= 0) {
      errors.push(`Row ${sheetRow}: Duration must be a positive number.`);
    }

    if (id) idSet.add(id);
    activities.push({ id, name, predecessors, duration, sourceRow: sheetRow });
  });

  activities.forEach(activity => {
    activity.predecessors.forEach(predecessor => {
      if (!idSet.has(predecessor)) {
        errors.push(`Row ${activity.sourceRow}: invalid predecessor "${predecessor}" for Activity No. "${activity.id}".`);
      }
      if (predecessor === activity.id) {
        errors.push(`Row ${activity.sourceRow}: activity cannot be its own predecessor.`);
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

  sched.getRange(SCHED_HEADER_ROW, 1, 1, 8).setValues([['Activity No.', 'Activity', 'Predecessor', 'Duration', 'Early Start', 'Early Finish', 'Late Start', 'Late Finish']]);
  sched.getRange(SCHED_FIRST_DATA_ROW, 1, output.length, 8).setValues(output);

  sched.getRange(SCHED_HEADER_ROW, GANTT_FIRST_COLUMN, 1, timeline.length).setValues([timeline]);

  const backgrounds = schedule.map(activity => timeline.map(day => {
    return day >= activity.earlyStart && day <= activity.earlyFinish ? '#4CAF50' : null;
  }));
  sched.getRange(SCHED_FIRST_DATA_ROW, GANTT_FIRST_COLUMN, backgrounds.length, timeline.length).setBackgrounds(backgrounds);
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
    const rangeMatch = token.match(/^(\d+)\s*-\s*(\d+)$/);

    if (rangeMatch) {
      expandPredecessorRange_(rangeMatch[1], rangeMatch[2]).forEach(id => predecessors.push(id));
    } else {
      predecessors.push(token);
    }
  });

  return Array.from(new Set(predecessors.filter(Boolean)));
}

function expandPredecessorRange_(startId, endId) {
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
