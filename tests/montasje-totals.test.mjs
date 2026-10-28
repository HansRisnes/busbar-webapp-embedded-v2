import assert from 'node:assert/strict';
import { test } from 'node:test';
import { getDashboardTotals } from '../src/client/dashboard.js';

test('montasje startup cost is included in cost without inflating profit', ()=>{
  const line = {
    selectedAddonConfig: { includeMontasje: true, includeEngineering: false, includeOppheng: false },
    totals: {
      totalExMontasje: 0,
      montasje: { cost: 1600 },
      montasjeStartupCost: 20000,
      montasjeMargin: 400,
      totalInclMontasje: 22000
    }
  };
  const project = { projectStatus: 'won', lines: [line] };
  const totals = getDashboardTotals([project]);

  assert.equal(totals.montasje.total, 22000);
  assert.equal(totals.montasje.cost, 21600);
  assert.equal(totals.montasje.margin, 400);
  assert.equal(totals.montasje.realProfit, 400);
  assert.equal(totals.allCost, 21600);
  assert.equal(totals.finishedAllProfit, 400);

  line.selectedAddonConfig.includeMontasje = false;
  const excluded = getDashboardTotals([project]);
  assert.equal(excluded.montasje.total, 0);
  assert.equal(excluded.montasje.cost, 0);
  assert.equal(excluded.allCost, 0);
});
