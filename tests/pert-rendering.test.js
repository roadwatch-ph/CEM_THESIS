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

didRenderFallback = false;
context.renderPertCompositeArrowImage_ = () => false;
context.renderPertImageArrow_ = () => true;
context.renderPertArrowGrid_ = () => { didRenderFallback = true; };

assert.strictEqual(
  context.renderPertArrows_({}, [], layout, 12, 12),
  true,
  'a successfully inserted image should still request node repainting for the durable connector layer'
);
assert.strictEqual(
  didRenderFallback,
  true,
  'a successfully inserted image must retain a visible cell connector backup'
);

const fallbackGrid = context.createPertArrowGrid_(20, 20);
context.drawPertSmartArrow_(fallbackGrid, position(0, 0), position(1, 4), 0, 0, 1, new Set());
const fallbackGlyphs = fallbackGrid.flat().filter(Boolean);
assert.ok(fallbackGlyphs.includes('▶'), 'the fallback must retain a visible arrowhead');
assert.ok(fallbackGlyphs.includes('━'), 'a downward dependency should include a continuous horizontal connector');
assert.ok(fallbackGlyphs.includes('┃'), 'a downward dependency should include a continuous vertical connector');
assert.ok(!fallbackGlyphs.includes('╲'), 'the fallback must not use detached diagonal slash glyphs');
assert.ok(!fallbackGlyphs.includes('╱'), 'the fallback must not use detached diagonal slash glyphs');

const horizontalGrid = context.createPertArrowGrid_(20, 20);
context.drawPertSmartArrow_(horizontalGrid, position(0, 0), position(1, 0), 0, 0, 1, new Set());
const horizontalGlyphs = horizontalGrid.flat().filter(Boolean);
assert.ok(horizontalGlyphs.includes('━'), 'a same-row dependency should use a horizontal fallback segment');
assert.ok(horizontalGlyphs.includes('▶'), 'a horizontal fallback must retain a visible arrowhead');

const verticalSource = position(2, 0);
const verticalTarget = position(2, 8);
const verticalPoints = context.getPertArrowPixelConnectionPoints_(verticalSource, verticalTarget, 0, 1, 0, 1);
assert.strictEqual(verticalPoints.start.x, verticalPoints.end.x, 'vertically aligned nodes should connect through top/bottom ports');
assert.ok(verticalPoints.start.y < verticalPoints.end.y, 'a downward dependency should point down');

const verticalRoute = context.getPertPreferredPixelRoutePoints_(verticalPoints.start, verticalPoints.end, 0, 0);
assert.strictEqual(verticalRoute.length, 2, 'a clear vertical dependency should be drawn as one straight line');
assert.strictEqual(verticalRoute[0].x, verticalRoute[1].x, 'the vertical line must not acquire a horizontal bend');

const diagonalPoints = context.getPertArrowPixelConnectionPoints_(position(0, 0), position(1, 5), 0, 1, 0, 1);
const diagonalRoute = context.getPertPreferredPixelRoutePoints_(diagonalPoints.start, diagonalPoints.end, 0, 0);
assert.strictEqual(diagonalRoute.length, 2, 'a clear diagonal dependency should be drawn as one straight line');
assert.notStrictEqual(diagonalRoute[0].x, diagonalRoute[1].x, 'a diagonal dependency needs horizontal movement');
assert.notStrictEqual(diagonalRoute[0].y, diagonalRoute[1].y, 'a diagonal dependency needs vertical movement');

const diagonalSvg = context.createPertArrowRouteSvg_(160, 160, diagonalRoute, '#123456');
assert.ok(diagonalSvg.includes('<polyline'), 'the arrow renderer should draw a continuous line');
assert.ok(diagonalSvg.includes('<polygon'), 'the arrow renderer should draw an arrowhead at the line end');
assert.ok(diagonalSvg.includes('#123456'), 'the arrow drawing should retain the route color');

console.log('PERT rendering fallback tests passed');
