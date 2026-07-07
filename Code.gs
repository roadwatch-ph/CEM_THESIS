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

  orderedActivities.forEach(activity => {
    const latestPredecessorFinish = activity.predecessors.reduce((latestFinish, predecessorId) => {
      return Math.max(latestFinish, scheduleById.get(predecessorId).finish);
    }, 0);

    const start = latestPredecessorFinish + 1;
    const finish = start + activity.duration - 1;

    scheduleById.set(activity.id, {
      id: activity.id,
      name: activity.name,
      predecessors: activity.predecessors,
      duration: activity.duration,
      start,
      finish,
      sourceRow: activity.sourceRow,
    });
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
    activity.start,
    activity.finish,
  ]);

  sched.getRange(1, 1, 1, 6).setValues([['Activity No.', 'Activity', 'Predecessor', 'Duration', 'Start Day', 'Finish Day']]);
  sched.getRange(2, 1, output.length, 6).setValues(output);

  const maxFinish = Math.max(...schedule.map(activity => activity.finish));
  const timeline = Array.from({ length: maxFinish }, (_, index) => index + 1);
  sched.getRange(1, 7, 1, timeline.length).setValues([timeline]);

  const backgrounds = schedule.map(activity => timeline.map(day => {
    return day >= activity.start && day <= activity.finish ? '#4CAF50' : null;
  }));
  sched.getRange(2, 7, backgrounds.length, timeline.length).setBackgrounds(backgrounds);
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

  return text
    .split(',')
    .map(predecessor => normalizeId_(predecessor))
    .filter(Boolean);
}
