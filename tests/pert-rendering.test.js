'use strict';

const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const source = fs.readFileSync('Code.gs', 'utf8');
const context = { console, Map, Set, Math, Array, String, Number, Object, RegExp, JSON };
vm.createContext(context);
vm.runInContext(source, context, { filename: 'Code.gs' });

function position(level, rowOffset) {
  return { level, renderedLevel: level, rowOffset };
}

const sourcePosition = position(0, 0);
const targetPosition = position(1, 0);
const layout = { positions: new Map([['START', sourcePosition], ['A', targetPosition]]) };
const routes = [{
  sourceId: 'START',
  targetId: 'A',
  sourcePosition,
  targetPosition,
  successorIndex: 0,
  successorCount: 1,
  incomingIndex: 0,
  incomingCount: 1,
  color: '#000000',
}];

assert.strictEqual(
  context.shouldRenderPertImageArrows_(Array.from({ length: 251 }), routes),
  false,
  'large diagrams should select the fallback instead of attempting over-grid images'
);

let didRenderFallback = false;
context.buildPertArrowRoutes_ = () => routes;
context.shouldRenderPertImageArrows_ = () => true;
context.renderPertCompositeArrowImage_ = () => false;
context.renderPertImageArrow_ = () => false;
context.renderPertArrowGrid_ = () => { didRenderFallback = true; };

assert.strictEqual(
  context.renderPertArrows_({}, [], layout, 12, 12),
  true,
  'a failed image route should request node repainting after the grid fallback'
);
assert.strictEqual(didRenderFallback, true, 'a failed image route should be rendered by the fallback');

const fallbackGrid = context.createPertArrowGrid_(20, 20);
context.drawPertSmartArrow_(fallbackGrid, position(0, 0), position(1, 4), 0, 0, 1, new Set());
const fallbackGlyphs = fallbackGrid.flat().filter(Boolean);
assert.ok(fallbackGlyphs.includes('▶'), 'the fallback must retain a visible arrowhead');
assert.ok(!fallbackGlyphs.includes('╲') && !fallbackGlyphs.includes('╱'), 'the fallback must not use visually disconnected diagonal glyphs');

console.log('PERT rendering fallback tests passed');
