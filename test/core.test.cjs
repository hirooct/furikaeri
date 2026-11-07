const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

test('児童分離、単元変更後の回答編集、教師2名の返信', () => {
  const tabs = new Map();
  let email = 'a@school.test';
  let sequence = 0;
  const book = {
    getId: () => 'book',
    getSheetByName: name => tabs.get(name),
    insertSheet(name) {
      const rows = [];
      const sheet = {
        rows, appendRow: row => rows.push(row), setFrozenRows() {},
        getLastRow: () => rows.length,
        getRange(row, col, count, width) {
          return {
            getDisplayValues: () => rows.slice(row - 1, row - 1 + count)
              .map(values => values.slice(col - 1, col - 1 + width).map(String)),
            setValues: values => rows.splice(row - 1, values.length, ...values)
          };
        }
      };
      tabs.set(name, sheet);
      return sheet;
    }
  };
  const props = new Map();
  const context = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => book, openById: () => book, flush() {} },
    Session: { getEffectiveUser: () => ({getEmail: () => 't@school.test'}),
      getActiveUser: () => ({getEmail: () => email}) },
    PropertiesService: { getScriptProperties: () => ({getProperty: key => props.get(key),
      setProperty: (key, value) => props.set(key, value)}) },
    LockService: {getScriptLock: () => ({waitLock() {}, releaseLock() {}})},
    Utilities: { getUuid: () => 'id-' + ++sequence }
  });
  vm.runInContext(fs.readFileSync(require.resolve('../src/Code.gs'), 'utf8'), context);
  context.setup_();
  const roster = tabs.get('名簿');
  roster.appendRow(['a@school.test', '児童A', 'student', 1]);
  roster.appendRow(['b@school.test', '児童B', 'student', 2]);
  roster.appendRow(['u@school.test', '共同担当', 'teacher', '']);
  const as = (address, action, data) => {
    email = address + '@school.test';
    return context.api(action, data);
  };
  const field = { type: 'text', label: '工夫したこと', required: true, examples: [], groups: [] };
  const setting = { subject: '体育', title: 'ハードル走', goal: '速くゴールする', hours: 4, fields: [field] };
  assert.throws(() => as('outsider', 'home'), /名簿/);
  const unit = as('t', 'saveUnit', setting);
  const selection = { unit: unit.id, period: 1, student: 'a@school.test' };
  const answer = { ...selection, revision: unit.revision, version: 0, answers: {[unit.fields[0].id]: '助走を合わせた'} };
  const record = as('a', 'save', answer);
  assert.equal(as('b', 'read', selection).length, 0);
  assert.throws(() => as('b', 'feedback', {...selection, record: record.id, stamp: '👍', comment: ''}), /利用できません/);
  assert.throws(() => as('a', 'save', answer), /更新/);
  as('t', 'saveUnit', { ...setting, id: unit.id, revision: unit.revision, fields: [{...field, label: '次の課題'}] });
  as('a', 'save', { ...answer, version: 1 });
  as('t', 'feedback', { ...selection, record: record.id, stamp: '👍', comment: 'いい工夫です' });
  as('u', 'feedback', { ...selection, record: record.id, stamp: '💡', comment: '' });
  const updated = as('a', 'read', selection)[0];
  assert.equal(updated.answers[0].label, '工夫したこと');
  assert.equal(updated.feedback.length, 2);
  const classView = as('u', 'read', { ...selection, student: '' });
  assert.equal(classView.length, 2);
  assert.equal(classView[1].record, null);
  assert.equal(as('u', 'aiPreview', selection).includes('児童A'), false);
});
