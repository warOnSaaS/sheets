// Example workbooks, all fictional. Each is built with the same operations the tools use, so a template
// is exactly what an agent could make itself. Used to seed the demo and by sheets.create_sheet's template.
import { Y, initBook, needTab, tabList } from './book.mjs';
import { Engine } from './engine.mjs';
import * as ops from './ops.mjs';

const me = { id: 'p_example', name: 'Sam Rivera', role: 'owner' };

export const TEMPLATES = {
  pipeline: { title: 'Acme Dental pipeline', about: 'Deals by stage with weighted value, a summary by stage and a chart' },
  budget: { title: 'Acme Dental 2026 budget', about: 'Monthly income and costs with totals, net and a chart' },
  hiring: { title: 'Hiring tracker', about: 'Candidates by role and stage, with a pivot table of the pipeline' },
};

export function buildTemplate(name) {
  const t = TEMPLATES[name];
  if (!t) return null;
  const doc = new Y.Doc();
  const build = { pipeline, budget, hiring }[name];
  initBook(doc, { title: t.title, tabs: build.tabs });
  const e = new Engine(doc);
  build(doc, e);
  e.destroy();
  return doc;
}

pipeline.tabs = [{ name: 'Deals' }, { name: 'Summary' }, { name: 'Stages' }];
function pipeline(doc, e) {
  ops.writeRange(doc, e, { tab: 'Stages', range: 'A1', values: [
    ['Stage', 'Probability'], ['Lead', '10%'], ['Qualified', '25%'], ['Proposal', '50%'], ['Negotiation', '75%'], ['Won', '100%'], ['Lost', '0%'],
  ] });
  ops.formatRange(doc, e, { tab: 'Stages', range: 'A1:B1', bold: true });
  ops.writeRange(doc, e, { tab: 'Deals', range: 'A1', values: [
    ['Deal', 'Company', 'Contact', 'Stage', 'Amount', 'Probability', 'Weighted', 'Close date', 'Owner'],
    ['Staff dental plan', 'Birch Law', 'Jordan Lee', 'Negotiation', '$18,000', null, null, '2026-10-16', 'Sam'],
    ['Family plan add-on', 'Harbor Fitness', 'Casey Morgan', 'Proposal', '$12,000', null, null, '2026-10-21', 'Sam'],
    ['Whitening for members', 'Harbor Fitness', 'Casey Morgan', 'Lead', '$3,500', null, null, '2026-11-30', 'Riley'],
    ['School screening day', 'Pinecrest School', 'Riley Chen', 'Qualified', '$6,200', null, null, '2026-11-12', 'Jordan'],
    ['Driver checkups', 'Summit Logistics', 'Avery Brooks', 'Proposal', '$22,400', null, null, '2026-11-03', 'Casey'],
    ['Staff dental plan', 'Northside Bakery', 'Morgan Patel', 'Won', '$4,800', null, null, '2026-09-17', 'Sam'],
    ['Clinic referrals', 'Lakeview Clinic', 'Quinn Alvarez', 'Negotiation', '$9,900', null, null, '2026-10-28', 'Riley'],
    ['Staff dental plan', 'Cedar Credit Union', 'Taylor Kim', 'Qualified', '$15,600', null, null, '2026-12-05', 'Jordan'],
    ['Orthodontics package', 'Riverbend Realty', 'Jamie Ortiz', 'Lost', '$7,300', null, null, '2026-09-02', 'Casey'],
    ['Staff dental plan', 'Oakridge Builders', 'Drew Nakamura', 'Lead', '$11,000', null, null, '2026-12-19', 'Sam'],
    ['Cleaning days', 'Maple Street Cafe', 'Parker Singh', 'Won', '$2,600', null, null, '2026-08-28', 'Riley'],
    ['Emergency care plan', 'Summit Logistics', 'Avery Brooks', 'Qualified', '$8,750', null, null, '2026-11-20', 'Casey'],
  ] });
  ops.setFormula(doc, e, { tab: 'Deals', cell: 'F2', formula: '=IFERROR(VLOOKUP(D2,Stages!$A$2:$B$7,2,FALSE),0)', fill: 'F2:F13' });
  ops.setFormula(doc, e, { tab: 'Deals', cell: 'G2', formula: '=E2*F2', fill: 'G2:G13' });
  ops.writeRange(doc, e, { tab: 'Deals', range: 'A15', values: [['Total', null, null, null, '=SUM(E2:E13)', null, '=SUM(G2:G13)']] });
  ops.formatRange(doc, e, { tab: 'Deals', range: 'A1:I1', bold: true, fill: '#f1f3f5' });
  ops.formatRange(doc, e, { tab: 'Deals', range: 'A15:I15', bold: true });
  ops.formatRange(doc, e, { tab: 'Deals', range: 'E2:E15', number_format: 'currency0' });
  ops.formatRange(doc, e, { tab: 'Deals', range: 'G2:G15', number_format: 'currency0' });
  ops.formatRange(doc, e, { tab: 'Deals', range: 'F2:F13', number_format: '0%' });
  ops.formatRange(doc, e, { tab: 'Deals', range: 'E1:G15', align: 'right' });
  ops.setFreeze(doc, { tab: 'Deals', rows: 1, columns: 1 });
  ops.setColumnWidth(doc, { tab: 'Deals', columns: 'A', width: 190 });
  ops.setColumnWidth(doc, { tab: 'Deals', columns: 'B:C', width: 150 });
  ops.setColumnWidth(doc, { tab: 'Deals', columns: 'D', width: 112 });

  ops.writeRange(doc, e, { tab: 'Summary', range: 'A1', values: [
    ['Stage', 'Deals', 'Amount', 'Weighted'],
    ...['Lead', 'Qualified', 'Proposal', 'Negotiation', 'Won', 'Lost'].map((s) => [s, null, null, null]),
    ['Total', '=SUM(B2:B7)', '=SUM(C2:C7)', '=SUM(D2:D7)'],
    [],
    ['Win rate', '=IFERROR(COUNTIF(Deals!D2:D13,"Won")/(COUNTIF(Deals!D2:D13,"Won")+COUNTIF(Deals!D2:D13,"Lost")),0)'],
    ['Open pipeline', '=SUMIFS(Deals!E2:E13,Deals!D2:D13,"<>Won",Deals!D2:D13,"<>Lost")'],
    ['Biggest open deal', '=MAXIFS(Deals!E2:E13,Deals!D2:D13,"<>Won",Deals!D2:D13,"<>Lost")'],
  ] });
  ops.setFormula(doc, e, { tab: 'Summary', cell: 'B2', formula: '=COUNTIF(Deals!$D$2:$D$13,A2)', fill: 'B2:B7' });
  ops.setFormula(doc, e, { tab: 'Summary', cell: 'C2', formula: '=SUMIF(Deals!$D$2:$D$13,A2,Deals!$E$2:$E$13)', fill: 'C2:C7' });
  ops.setFormula(doc, e, { tab: 'Summary', cell: 'D2', formula: '=SUMIF(Deals!$D$2:$D$13,A2,Deals!$G$2:$G$13)', fill: 'D2:D7' });
  ops.formatRange(doc, e, { tab: 'Summary', range: 'A1:D1', bold: true, fill: '#f1f3f5' });
  ops.formatRange(doc, e, { tab: 'Summary', range: 'A8:D8', bold: true });
  ops.formatRange(doc, e, { tab: 'Summary', range: 'C2:D8', number_format: 'currency0' });
  ops.formatRange(doc, e, { tab: 'Summary', range: 'B10', number_format: 'percent' });
  ops.formatRange(doc, e, { tab: 'Summary', range: 'B11:B12', number_format: 'currency0' });
  ops.formatRange(doc, e, { tab: 'Summary', range: 'A10:A12', bold: true });
  ops.setColumnWidth(doc, { tab: 'Summary', columns: 'A', width: 150 });
  ops.createChart(doc, e, { tab: 'Summary', type: 'column', range: 'A1:D7', series: ['C', 'D'], title: 'Pipeline by stage', at: 'F1' });
  ops.addComment(doc, { id: 'p_example_jordan', name: 'Jordan Lee', role: 'member' }, { tab: 'Deals', cell: 'E2', body: 'Their office manager asked for a second quote with orthodontics included. Holding at $18k until we hear back.' });
  ops.addComment(doc, me, { tab: 'Summary', cell: 'B10', body: 'Win rate counts only closed deals (won and lost).' });
}

budget.tabs = [{ name: 'Monthly' }, { name: 'Assumptions' }];
function budget(doc, e) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  ops.writeRange(doc, e, { tab: 'Assumptions', range: 'A1', values: [
    ['Assumption', 'Value'], ['Patients per month', 420], ['Average visit', '$185'], ['Monthly growth', '1.5%'], ['Rent', '$6,800'], ['Staff', '$38,500'], ['Supplies, share of income', '9%'],
  ] });
  ops.formatRange(doc, e, { tab: 'Assumptions', range: 'A1:B1', bold: true });
  ops.setColumnWidth(doc, { tab: 'Assumptions', columns: 'A', width: 210 });
  ops.writeRange(doc, e, { tab: 'Monthly', range: 'A1', values: [['Line', ...months, 'Total', 'Average']] });
  const rows = [['Patient visits'], ['Insurance plans'], ['Income'], [], ['Rent'], ['Staff'], ['Supplies'], ['Marketing'], ['Costs'], [], ['Net'], ['Margin']];
  ops.writeRange(doc, e, { tab: 'Monthly', range: 'A2', values: rows });
  ops.setFormula(doc, e, { tab: 'Monthly', cell: 'B2', formula: '=ROUND(Assumptions!$B$2*Assumptions!$B$3*(1+Assumptions!$B$4)^(COLUMN()-2),0)', fill: 'B2:M2' });
  ops.writeRange(doc, e, { tab: 'Monthly', range: 'B3', values: [[9200, 9200, 9400, 9400, 9650, 9650, 9900, 9900, 10150, 10150, 10400, 10400]] });
  ops.setFormula(doc, e, { tab: 'Monthly', cell: 'B4', formula: '=B2+B3', fill: 'B4:M4' });
  ops.setFormula(doc, e, { tab: 'Monthly', cell: 'B6', formula: '=Assumptions!$B$5', fill: 'B6:M6' });
  ops.setFormula(doc, e, { tab: 'Monthly', cell: 'B7', formula: '=Assumptions!$B$6', fill: 'B7:M7' });
  ops.setFormula(doc, e, { tab: 'Monthly', cell: 'B8', formula: '=ROUND(B4*Assumptions!$B$7,0)', fill: 'B8:M8' });
  ops.writeRange(doc, e, { tab: 'Monthly', range: 'B9', values: [[2400, 1800, 1800, 2400, 1800, 1800, 1200, 1200, 3000, 2400, 1800, 1200]] });
  ops.setFormula(doc, e, { tab: 'Monthly', cell: 'B10', formula: '=SUM(B6:B9)', fill: 'B10:M10' });
  ops.setFormula(doc, e, { tab: 'Monthly', cell: 'B12', formula: '=B4-B10', fill: 'B12:M12' });
  ops.setFormula(doc, e, { tab: 'Monthly', cell: 'B13', formula: '=IFERROR(B12/B4,0)', fill: 'B13:M13' });
  for (const r of [2, 3, 4, 6, 7, 8, 9, 10, 12]) {
    ops.setFormula(doc, e, { tab: 'Monthly', cell: `N${r}`, formula: `=SUM(B${r}:M${r})` });
    ops.setFormula(doc, e, { tab: 'Monthly', cell: `O${r}`, formula: `=AVERAGE(B${r}:M${r})` });
  }
  ops.setFormula(doc, e, { tab: 'Monthly', cell: 'N13', formula: '=IFERROR(N12/N4,0)' });
  ops.formatRange(doc, e, { tab: 'Monthly', range: 'B2:O12', number_format: 'currency0' });
  ops.formatRange(doc, e, { tab: 'Monthly', range: 'B13:O13', number_format: 'percent' });
  ops.formatRange(doc, e, { tab: 'Monthly', range: 'A1:O1', bold: true, fill: '#f1f3f5' });
  ops.formatRange(doc, e, { tab: 'Monthly', range: 'A4:O4', bold: true });
  ops.formatRange(doc, e, { tab: 'Monthly', range: 'A10:O10', bold: true });
  ops.formatRange(doc, e, { tab: 'Monthly', range: 'A12:O12', bold: true, color: '#1a7f37' });
  ops.formatRange(doc, e, { tab: 'Monthly', range: 'B1:O13', align: 'right' });
  ops.setFreeze(doc, { tab: 'Monthly', rows: 1, columns: 1 });
  ops.setColumnWidth(doc, { tab: 'Monthly', columns: 'A', width: 150 });
  ops.setColumnWidth(doc, { tab: 'Monthly', columns: 'B:O', width: 92 });
  const ch = ops.createChart(doc, e, { tab: 'Monthly', type: 'line', range: 'A1:M13', series_by: 'rows', series: [4, 10, 12], title: 'Income, costs and net by month', at: 'A16' });
  ops.updateChart(doc, e, { chart: ch.id, width: 760, height: 300 });
  ops.addComment(doc, me, { tab: 'Monthly', cell: 'J9', body: 'September is higher: the back-to-school mailer.' });
}

hiring.tabs = [{ name: 'Candidates' }];
function hiring(doc, e) {
  ops.writeRange(doc, e, { tab: 'Candidates', range: 'A1', values: [
    ['Candidate', 'Role', 'Stage', 'Applied', 'Interviewer', 'Score', 'Next step'],
    ['Avery Brooks', 'Dental hygienist', 'Interview', '2026-09-14', 'Jordan', 4.5, 'Second interview Thu'],
    ['Morgan Patel', 'Front desk', 'Offer', '2026-09-02', 'Casey', 4.8, 'Send offer letter'],
    ['Quinn Alvarez', 'Dental hygienist', 'Screen', '2026-09-28', 'Jordan', 3.9, 'Phone screen'],
    ['Taylor Kim', 'Dental assistant', 'Interview', '2026-09-21', 'Riley', 4.1, 'Skills check'],
    ['Jamie Ortiz', 'Front desk', 'Rejected', '2026-08-30', 'Casey', 2.7, null],
    ['Drew Nakamura', 'Dental assistant', 'Applied', '2026-10-03', null, null, 'Review application'],
    ['Reese Johnson', 'Dental hygienist', 'Applied', '2026-10-05', null, null, 'Review application'],
    ['Parker Singh', 'Office manager', 'Screen', '2026-09-30', 'Sam', 4.0, 'Phone screen'],
    ['Rowan Lee', 'Dental assistant', 'Hired', '2026-08-12', 'Riley', 4.6, 'Starts Oct 13'],
    ['Emerson Diaz', 'Office manager', 'Interview', '2026-09-25', 'Sam', 4.3, 'Meet the team'],
  ] });
  ops.writeRange(doc, e, { tab: 'Candidates', range: 'H1', values: [['Days since applied']] });
  ops.setFormula(doc, e, { tab: 'Candidates', cell: 'H2', formula: '=IF(OR(C2="Hired",C2="Rejected"),"",DATE(2026,10,7)-D2)', fill: 'H2:H11' });
  ops.formatRange(doc, e, { tab: 'Candidates', range: 'A1:H1', bold: true, fill: '#f1f3f5' });
  ops.formatRange(doc, e, { tab: 'Candidates', range: 'F2:F11', number_format: '0.0' });
  ops.formatRange(doc, e, { tab: 'Candidates', range: 'H2:H11', number_format: '0' });
  ops.formatRange(doc, e, { tab: 'Candidates', range: 'C2:C11', bold: true });
  ops.setFreeze(doc, { tab: 'Candidates', rows: 1 });
  ops.setColumnWidth(doc, { tab: 'Candidates', columns: 'A:B', width: 150 });
  ops.setColumnWidth(doc, { tab: 'Candidates', columns: 'G', width: 180 });
  ops.setColumnWidth(doc, { tab: 'Candidates', columns: 'H', width: 140 });
  ops.createPivot(doc, e, { tab: 'Candidates', range: 'A1:H11', rows: 'Role', values: [{ summarize: 'count' }, { field: 'Score', summarize: 'average' }], target: 'By role' });
  ops.formatRange(doc, e, { tab: 'By role', range: 'C2:C6', number_format: '0.0' });
  ops.setColumnWidth(doc, { tab: 'By role', columns: 'A', width: 160 });
  ops.setColumnWidth(doc, { tab: 'By role', columns: 'B:C', width: 140 });
  ops.createChart(doc, e, { tab: 'By role', type: 'bar', range: 'A1:B5', title: 'Candidates by role', at: 'E1' });
  ops.addComment(doc, me, { tab: 'Candidates', cell: 'C3', body: 'Morgan accepted verbally. Offer letter goes out Monday.' });
}

export const templateTabs = (doc) => tabList(doc).map((t) => t.name);
export { needTab };
