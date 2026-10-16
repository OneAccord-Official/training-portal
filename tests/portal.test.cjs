const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');

class Element {
  constructor() { this.children = []; this.listeners = {}; this.attributes = {}; }
  appendChild(el) { this.children.push(el); }
  setAttribute(name, value) { this.attributes[name] = value; }
  addEventListener(name, fn) { this.listeners[name] = fn; }
}
function harness(fetch) {
  const stream = new Element();
  const context = vm.createContext({
    fetch, console: { warn() {}, error() {} },
    document: { getElementById: () => stream, createElement: () => new Element() },
  });
  const source = fs.readFileSync(path.join(root, 'js/app.js'), 'utf8').replace('\ninit();', '\n');
  vm.runInContext(source + '\nrenderNav = () => {}; globalThis.api = {state, recordCompletion, renderLessonComplete};', context);
  context.api.state.personEmail = 'learner@example.org';
  return { ...context.api, stream };
}
test('all indexed ready lessons exist, parse, and use supported blocks', () => {
  const catalog = JSON.parse(fs.readFileSync(path.join(root, 'data/trainings-index.json')));
  const trainingIds = new Set();
  const supported = new Set(['text','video','activity','question','links','scenario','deliverable','reflection','gameplan']);
  for (const training of catalog.trainings) {
    assert.ok(!trainingIds.has(training.id), `Duplicate training: ${training.id}`);
    trainingIds.add(training.id);
    const ids = new Set();
    for (const lesson of training.lessons) {
      assert.ok(!ids.has(lesson.id), `Duplicate lesson: ${training.id}/${lesson.id}`);
      ids.add(lesson.id);
      if (lesson.status !== 'ready') continue;
      const data = JSON.parse(fs.readFileSync(path.join(root, 'data/trainings', training.id, `${lesson.id}.json`)));
      assert.equal(data.id, lesson.id);
      assert.ok(data.blocks.length > 0);
      for (const block of data.blocks) assert.ok(supported.has(block.type), `Unknown block: ${block.type}`);
      if (data.nextLessonId) assert.ok(training.lessons.some(l => l.id === data.nextLessonId), 'Broken next lesson link');
    }
  }
});
test('failed completion write does not create a checkmark', async () => {
  const h = harness(async () => ({ ok: false, status: 404 }));
  assert.equal(await h.recordCompletion('course', 'lesson'), false);
  assert.equal(h.state.completions.size, 0);
});
test('network failure is retryable, not success', async () => {
  const h = harness(async () => { throw new Error('Offline'); });
  assert.equal(await h.recordCompletion('course', 'lesson'), false);
  assert.equal(h.state.completions.size, 0);
});
test('success requires explicit server acknowledgement', async () => {
  const h = harness(async () => ({ ok: true, json: async () => ({}) }));
  assert.equal(await h.recordCompletion('course', 'lesson'), false);
  assert.equal(h.state.completions.size, 0);
});
test('confirmed server write creates the correct checkmark', async () => {
  let payload;
  const h = harness(async (_, opts) => {
    payload = JSON.parse(opts.body);
    return { ok: true, json: async () => ({ok: true, id: 'test-item'}) };
  });
  assert.equal(await h.recordCompletion('course', 'lesson'), true);
  assert.equal(h.state.completions.has('course:lesson'), true);
  assert.equal(payload.personEmail, 'learner@example.org');
});
test('no learner identity never submits a write', async () => {
  const h = harness(async () => { throw new Error('Must not fetch'); });
  h.state.personEmail = null;
  assert.equal(await h.recordCompletion('course', 'lesson'), false);
});
test('completion UI reports failure and retry uses the original lesson after navigation', async () => {
  let attempts = 0;
  const payloads = [];
  const h = harness(async (_, opts) => {
    payloads.push(JSON.parse(opts.body));
    attempts++;
    return attempts === 1 ? {ok: false, status: 503} : {ok: true, json: async () => ({ok: true})};
  });
  h.state.currentTrainingId = 'original';
  h.state.currentLessonId = 'first';
  h.state.lessonMeta = {};
  h.renderLessonComplete();
  await new Promise(resolve => setImmediate(resolve));
  const complete = h.stream.children[0];
  const [status, retry] = complete.children;
  assert.match(status.textContent, /NOT been saved/);
  assert.equal(retry.hidden, false);
  h.state.currentTrainingId = 'other';
  h.state.currentLessonId = 'second';
  await retry.listeners.click();
  assert.match(status.textContent, /Completion saved/);
  assert.equal(retry.hidden, true);
  assert.equal(payloads[1].trainingId, 'original');
  assert.equal(payloads[1].lessonId, 'first');
});
