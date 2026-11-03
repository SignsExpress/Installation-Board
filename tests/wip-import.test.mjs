import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import xlsx from 'xlsx';

function importer(path) {
  const source = fs.readFileSync(path, 'utf8');
  const context = vm.createContext({});
  const aliases = source.match(/const WIP_HEADER_ALIASES = \{[\s\S]*?\n\};/)[0];
  const names = ['normalizeWipKey', 'normalizeWipOrderReference', 'getWipLocation',
    'getWipSalesperson', 'formatWipMoney', 'getWipTabFromStatus', 'getWipHeaderIndex',
    'isWipOrderHeaderLabel', 'normalizeWipAssignees', 'makeWipCard', 'parseWipWorkbook',
    'isWipPlaceholderDescription', 'isWipCardInCompletedLane'];
  const functions = names.map(name => {
    const start = source.indexOf('function ' + name + '(');
    return source.slice(start, source.indexOf('\n}', start) + 2);
  });
  vm.runInContext([aliases, ...functions].join('\n'), context);
  return (buffer, cards = []) => context.parseWipWorkbook(buffer, cards, xlsx);
}

for (const path of ['src/App.jsx', 'src/ModernApp.jsx']) {
  test(`${path}: imports WIP by Order CSV and retains remembered planning`, () => {
    const csv = '"Order #","Status","Customer","Salesperson","Order Due Date","Status Start Date","Order Subtotal","Location","Number of Line Items on Orders","On Hold","Days in Status","Overdue Production","Description"\r\n'
      + '"ORD-101","Pre-WIP","Customer, Ltd","Tom Example","25 August 2026","01 January 1901","13333.8","Signs Express (Central Lancashire)","13","No","","Yes","Signs"\r\n'
      + '"ORDS-102","WIP","Other customer","Amber Example","25 August 2026","01 January 1901","250","Signs Express Southport","2","No","","Yes","Door graphics"';
    const cards = importer(path)(Buffer.from(csv), [{orderNumber: 'ORDS-102', id: 'saved', lane: 'ordered', productionAssignees: ['Tom']}]);
    assert.equal(cards.length, 2);
    assert.equal(cards[0].company, 'Customer, Ltd');
    assert.equal(cards[0].tab, 'pre-wip');
    assert.equal(cards[0].preTaxTotal, '£13333.8');
    assert.equal(cards[0].productionLocation, 'Central Lancs');
    assert.equal(cards[1].productionLocation, 'Southport');
    assert.equal(cards[1].description, 'Door graphics');
    assert.equal(cards[1].lane, 'ordered');
    assert.equal(cards[1].id, 'saved');
    assert.equal(cards[1].productionAssignees[0], 'Tom');
  });

  test(`${path}: preserves Excel headers and rejects missing required columns`, () => {
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, xlsx.utils.aoa_to_sheet([
      ['Order Number', 'Order Status', 'Company', 'Production Location', 'Pre-Tax Total'],
      ['ORD-103', 'WIP', 'Customer', 'Southport', 125]
    ]), 'WIP');
    const parse = importer(path);
    const [card] = parse(xlsx.write(workbook, {type: 'buffer', bookType: 'xlsx'}));
    assert.equal(card.preTaxTotal, '£125');
    assert.equal(card.productionLocation, 'Southport');
    assert.throws(() => parse(Buffer.from('Order #,Status,Customer\nORD-103,WIP,Customer')), /Missing columns/);
  });
}
