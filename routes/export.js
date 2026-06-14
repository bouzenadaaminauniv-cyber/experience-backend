import express from 'express'
import { supabase } from '../lib/supabase.js'
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, VerticalAlign, ShadingType,
  PageOrientation, HeightRule, UnderlineType
} from 'docx'

const router = express.Router()

const YELLOW = 'F5E642'
const BLACK = '000000'
const WHITE = 'FFFFFF'
const TASKS_PER_PAGE = 10

function bd(size = 4) {
  return { style: BorderStyle.SINGLE, size, color: BLACK }
}
function allBorders(size = 4) {
  const b = bd(size)
  return { top: b, bottom: b, left: b, right: b }
}

function makeCell(text, opts = {}) {
  const {
    bold = false, size = 14, width, colspan = 1, rowspan = 1,
    align = AlignmentType.CENTER, valign = VerticalAlign.CENTER,
    fill = WHITE, italics = false, font = 'Arial'
  } = opts
  return new TableCell({
    columnSpan: colspan, rowSpan: rowspan,
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    verticalAlign: valign,
    borders: allBorders(),
    shading: { fill, type: ShadingType.CLEAR, color: 'auto' },
    margins: { top: 30, bottom: 30, left: 60, right: 60 },
    children: [new Paragraph({
      alignment: align,
      spacing: { line: 240 },
      children: [new TextRun({ text: String(text ?? ''), bold, size, italics, font, color: BLACK })]
    })]
  })
}

function chk(val) { return val ? 'X' : '' }

async function getLogoUrl() {
  const { data } = await supabase.from('settings').select('value').eq('key', 'logo_url').single()
  return data?.value || ''
}

function emptyRow() {
  return `<tr class="empty-row">${Array(23).fill('<td></td>').join('')}</tr>`
}

function padToTen(rowsHtml, count) {
  const padding = TASKS_PER_PAGE - count
  return rowsHtml + Array(Math.max(0, padding)).fill(emptyRow()).join('')
}

const TABLE_COLS = ['Date','Location','A/C or Comp. Type','A/C Reg. or Comp. PN / SN',
  'Type of Maintenance (Rating)','Privilege Used',
  'FOT','SGH','R/I','MEL','TS','MOD','REP','INSP',
  'Training','Perform','Supervise','CRS',
  'A T A','Operation Performed','Time Duration',
  'Maintenance record Ref.','Remarks']

const COLGROUP = `<colgroup>
  <col style="width:4.2%"><col style="width:4.2%">
  <col style="width:5.8%"><col style="width:5.8%">
  <col style="width:5.2%"><col style="width:3.2%">
  <col style="width:1.9%"><col style="width:1.9%"><col style="width:1.9%"><col style="width:1.9%">
  <col style="width:1.9%"><col style="width:1.9%"><col style="width:1.9%"><col style="width:1.9%">
  <col style="width:2.4%"><col style="width:2.4%"><col style="width:2.4%"><col style="width:2.4%">
  <col style="width:3.8%"><col style="width:11%">
  <col style="width:3.8%"><col style="width:11%"><col style="width:5.5%">
</colgroup>`

const TABLE_THEAD = `<thead>
  <tr>
    <th class="num-th" rowspan="2">1.</th><th class="num-th" rowspan="2">2.</th>
    <th class="num-th" rowspan="2">3.</th><th class="num-th" rowspan="2">4.</th>
    <th class="num-th" rowspan="2">5.</th><th class="num-th" rowspan="2">6.</th>
    <th class="num-th" colspan="8">7.</th>
    <th class="num-th" colspan="4">8.</th>
    <th class="num-th" rowspan="2">9.</th><th class="num-th" rowspan="2">10.</th>
    <th class="num-th" rowspan="2">11.</th><th class="num-th" rowspan="2">12.</th>
    <th class="num-th" rowspan="2">13.</th>
  </tr>
  <tr>
    <th colspan="8" style="font-size:7.5px;background:#F5E642 !important;-webkit-print-color-adjust:exact;print-color-adjust:exact">Task Type</th>
    <th colspan="4" style="font-size:7.5px;background:#F5E642 !important;-webkit-print-color-adjust:exact;print-color-adjust:exact">Type of Activity</th>
  </tr>
  <tr>
    ${TABLE_COLS.map(l => `<th><span class="col-label">${l}</span></th>`).join('')}
  </tr>
</thead>`

const CSS = `
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:Arial,sans-serif; font-size:8.5px; background:white; }

  .page {
    width:287mm;
    height:200mm;
    padding:5mm 6mm;
    display:flex;
    flex-direction:column;
    overflow:hidden;
    page-break-after:always;
  }
  .page:last-child { page-break-after:auto; }

  .doc-table { width:100%; border-collapse:collapse; margin-bottom:2px; table-layout:fixed; }
  .doc-table td, .doc-table th { border:1px solid #000; padding:1px 3px; vertical-align:middle; overflow:hidden; word-wrap:break-word; }

  .header-table { table-layout:auto; margin-bottom:3px; }
  .header-logo { width:150px; vertical-align:top; padding:4px 6px !important; }
  .logo-img { width:105px; display:block; }
  .logo-technics { font-size:9.5px; font-weight:bold; margin-top:2px; }
  .logo-ref { font-size:7px; color:#444; }
  .header-title { text-align:center; vertical-align:middle; padding:5px !important; }
  .title-main { font-size:15px; font-weight:bold; letter-spacing:0.5px; }
  .title-sub { font-size:11px; font-style:italic; margin-top:3px; }
  .header-ids { width:145px; font-size:9.5px; vertical-align:top; padding:5px !important; line-height:2; }

  .records-banner { display:flex; justify-content:space-between; align-items:center; font-size:7.5px; padding:2px 0; margin-bottom:2px; }
  .records-note { max-width:240px; color:#333; font-style:italic; }
  .records-note-right { text-align:right; }
  .records-center { text-align:center; }
  .records-center strong { font-size:11px; display:block; }
  .records-center em { font-size:8px; }

  .main-table { font-size:7.5px; flex:1; }
  .main-table th {
    background:#F5E642 !important;
    -webkit-print-color-adjust:exact;
    print-color-adjust:exact;
    font-weight:bold; text-align:center;
    vertical-align:bottom; border:1px solid #000; padding:1px;
  }
  .num-th { font-size:9px; font-weight:bold; vertical-align:top !important; padding:1px; background:#F5E642 !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
  .col-label { display:block; font-size:7px; font-weight:normal; writing-mode:vertical-rl; transform:rotate(180deg); height:48px; line-height:1.1; margin:0 auto; }
  .main-table td { font-size:7.5px; text-align:center; padding:1px 2px; border:1px solid #000; vertical-align:middle; }
  .empty-row td { background:#fafafa; }
  .op-cell { text-align:left !important; font-size:7px; }
  .ref-cell { text-align:left !important; font-size:7px; }
  .check { font-weight:bold; font-size:9px; }

  .personnel-title { text-align:center; padding:16px 0 12px; font-size:12px; line-height:1.8; }
  .personnel-table { width:80%; margin:0 auto; border-collapse:collapse; }
  .personnel-table td { border:1px solid #000; padding:6px 12px; font-size:9.5px; vertical-align:middle; }
  .pers-label { width:190px; text-align:right; background:#fafafa; line-height:1.5; }

  .doc-note { font-size:7px; font-style:italic; text-align:center; padding:2px 0; border-top:1px solid #000; border-bottom:1px solid #000; margin:1px 0 2px; }
  .sig-block { display:flex; border:1px solid #000; margin-bottom:2px; }
  .sig-col { flex:1; padding:4px 7px; font-size:8px; line-height:1.5; border-right:1px solid #000; }
  .sig-col:last-child { border-right:none; }
  .sig-fields { display:flex; gap:8px; margin-top:3px; font-size:7.5px; }
  .sig-fields span { flex:1; }
  .doc-declaration { font-size:7px; font-style:italic; text-align:center; padding:2px 0; margin-bottom:1px; }
  .footer-ref { display:flex; justify-content:space-between; font-size:7px; font-style:italic; color:#555; margin-top:auto; padding-top:3px; }
`

function buildPageHtml(headerHtml, bodyHtml, sigName = '') {
  return `
  <div class="page">
    ${headerHtml}
    <div class="records-banner">
      <div class="records-note"><em>This logbook is intended to be hand written or filled electronically.<br>Add / print pages as needed.</em></div>
      <div class="records-center"><strong>RECORDS</strong><em>ENREGISTREMENTS</em></div>
      <div class="records-note records-note-right"><em>Ce livret est destiné à être renseigné à la main ou rempli électroniquement.<br>Ajouter / imprimer les pages au besoin.</em></div>
    </div>
    <table class="doc-table main-table" style="table-layout:fixed">
      ${COLGROUP}
      ${TABLE_THEAD}
      <tbody>${bodyHtml}</tbody>
    </table>
    <div class="doc-note">
      <strong>Note :</strong> The tasks recorded need to be related to a combination of activities / authorizations and not limited to simple tasks. &nbsp;/&nbsp;
      <strong>Note :</strong> Les tâches enregistrées doivent être associées à une combinaison d'activités / autorisations et ne se limitent pas à des tâches simples.
    </div>
    <div class="sig-block">
      <div class="sig-col">
        <strong>Logbook owner's signature (*) :</strong><br>
        <em>Signature du détenteur du livret</em>
        <div class="sig-fields">
          <span>Date :<br>…………………………….</span>
          <span>Name (Nom) :<br><strong>${sigName}</strong></span>
          <span>Signature :<br>…………………………….</span>
        </div>
      </div>
      <div class="sig-col">
        <strong>Team Leader / Manager's validation :</strong><br>
        <em>Validation du Chef d'Equipe / Responsable</em>
        <div class="sig-fields">
          <span>Date :<br>…………………………….</span>
          <span>Name (Nom) :<br>…………………………….</span>
          <span>Signature :<br>…………………………….</span>
        </div>
      </div>
    </div>
    <div class="doc-declaration">
      <em>(*) I declare that the entries in this logbook are complete and true. / Je <u>déclare</u> que les <u>enregistrements</u> de ce <u>livret</u> sont <u>complets</u> et <u>vrais</u>.</em>
    </div>
    <div class="footer-ref"><span>AHDT - 101</span><span>- Sept. 2018 -</span></div>
  </div>`
}

// ── BUILD WORD DOC ────────────────────────────────────────────────────────────
function buildDoc(entry, protocol, tasks, rowMap, user) {
  const aircraft = entry.ac_type || protocol?.aircraft_types?.name || ''
  const rating = entry.rating || protocol?.rating || ''
  const privilege = entry.privilege || protocol?.privilege || ''
  const PW = 15300

  const CW = {
    date: 580, loc: 580, actype: 900, acreg: 900,
    rating: 900, priv: 580,
    tt: 290, act: 400,
    ata: 580, op: 1600, dur: 580, ref: 1600, rem: 1080
  }
  const totalW = CW.date + CW.loc + CW.actype + CW.acreg + CW.rating + CW.priv
    + CW.tt * 8 + CW.act * 4 + CW.ata + CW.op + CW.dur + CW.ref + CW.rem
  const scale = PW / totalW
  const S = {}
  Object.keys(CW).forEach(k => { S[k] = Math.round(CW[k] * scale) })

  const allColWidths = [
    S.date, S.loc, S.actype, S.acreg, S.rating, S.priv,
    ...Array(8).fill(S.tt), ...Array(4).fill(S.act),
    S.ata, S.op, S.dur, S.ref, S.rem
  ]

  const headerTable = new Table({
    width: { size: PW, type: WidthType.DXA },
    columnWidths: [Math.round(PW * 0.15), Math.round(PW * 0.60), Math.round(PW * 0.25)],
    rows: [new TableRow({
      height: { value: 700, rule: HeightRule.ATLEAST },
      children: [
        new TableCell({
          borders: allBorders(),
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [
            new Paragraph({ children: [new TextRun({ text: 'AIR ALGÉRIE', bold: true, size: 18, font: 'Arial', color: 'CC0000' })] }),
            new Paragraph({ children: [new TextRun({ text: 'TECHNICS', bold: true, size: 16, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: 'DACM.AMO.01 / EASA.145.0163', size: 12, font: 'Arial' })] }),
          ]
        }),
        new TableCell({
          verticalAlign: VerticalAlign.CENTER,
          borders: allBorders(),
          margins: { top: 30, bottom: 30, left: 60, right: 60 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'MAINTENANCE EXPERIENCE LOGBOOK', bold: true, size: 22, font: 'Arial' })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "LIVRET D'EXPERIENCE EN MAINTENANCE", italics: true, size: 16, font: 'Arial' })] }),
          ]
        }),
        new TableCell({
          verticalAlign: VerticalAlign.CENTER,
          borders: allBorders(),
          margins: { top: 30, bottom: 30, left: 100, right: 60 },
          children: [
            new Paragraph({ children: [new TextRun({ text: `AHT ID :   ${user?.aht_id || '………………'}`, size: 14, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: `Page :      ………………`, size: 14, font: 'Arial' })] }),
          ]
        })
      ]
    })]
  })

  const recordsPara = new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80 },
    children: [
      new TextRun({ text: 'RECORDS', bold: true, size: 20, font: 'Arial' }),
      new TextRun({ text: '   ENREGISTREMENTS', italics: true, size: 14, font: 'Arial' }),
    ]
  })

  const leftNote = new Paragraph({
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text: 'This logbook is intended to be hand written or filled electronically. Add / print pages as needed.', italics: true, size: 11, font: 'Arial' })]
  })

  const hdr1 = new TableRow({
    height: { value: 700, rule: HeightRule.ATLEAST },
    children: [
      makeCell('1.\nDate', { bold: true, size: 11, fill: YELLOW }),
      makeCell('2.\nLocation', { bold: true, size: 11, fill: YELLOW }),
      makeCell('3.\nA/C or\nComp. Type', { bold: true, size: 11, fill: YELLOW }),
      makeCell('4.\nA/C Reg.\nor Comp.\nPN / SN', { bold: true, size: 11, fill: YELLOW }),
      makeCell('5.\nType of\nMaintenance\n(Rating)', { bold: true, size: 11, fill: YELLOW }),
      makeCell('6.\nPrivilege\nUsed', { bold: true, size: 11, fill: YELLOW }),
      makeCell('7.\nTask Type', { bold: true, size: 11, colspan: 8, fill: YELLOW }),
      makeCell('8.\nType of\nActivity', { bold: true, size: 11, colspan: 4, fill: YELLOW }),
      makeCell('9.\nATA', { bold: true, size: 11, fill: YELLOW }),
      makeCell('10.\nOperation\nPerformed', { bold: true, size: 11, fill: YELLOW }),
      makeCell('11.\nTime\nDuration', { bold: true, size: 11, fill: YELLOW }),
      makeCell('12.\nMaintenance\nRecord Ref.', { bold: true, size: 11, fill: YELLOW }),
      makeCell('13.\nRemarks', { bold: true, size: 11, fill: YELLOW }),
    ]
  })

  const hdr2 = new TableRow({
    height: { value: 500, rule: HeightRule.ATLEAST },
    children: [
      makeCell('', { fill: YELLOW }), makeCell('', { fill: YELLOW }),
      makeCell('', { fill: YELLOW }), makeCell('', { fill: YELLOW }),
      makeCell('', { fill: YELLOW }), makeCell('', { fill: YELLOW }),
      ...['FOT','SGH','R/I','MEL','TS','MOD','REP','INSP'].map(tt => makeCell(tt, { bold: true, size: 10, fill: YELLOW })),
      ...['Training','Perform','Supervise','CRS'].map(a => makeCell(a, { bold: true, size: 10, fill: YELLOW })),
      makeCell('', { fill: YELLOW }), makeCell('', { fill: YELLOW }),
      makeCell('', { fill: YELLOW }), makeCell('', { fill: YELLOW }),
      makeCell('', { fill: YELLOW }),
    ]
  })

  const taskRows = tasks.map(t => {
    const r = rowMap[t.id] || {}
    return new TableRow({
      height: { value: 360, rule: HeightRule.ATLEAST },
      children: [
        makeCell(entry.date || '', { size: 12 }),
        makeCell(entry.location || '', { size: 12 }),
        makeCell(aircraft, { size: 12 }),
        makeCell(entry.ac_registration || '', { size: 12 }),
        makeCell(rating, { size: 12 }),
        makeCell(privilege, { size: 12 }),
        makeCell(chk(r.task_type_fot), { size: 14, bold: true }),
        makeCell(chk(r.task_type_sgh), { size: 14, bold: true }),
        makeCell(chk(r.task_type_ri), { size: 14, bold: true }),
        makeCell(chk(r.task_type_mel), { size: 14, bold: true }),
        makeCell(chk(r.task_type_ts), { size: 14, bold: true }),
        makeCell(chk(r.task_type_mod), { size: 14, bold: true }),
        makeCell(chk(r.task_type_rep), { size: 14, bold: true }),
        makeCell(chk(r.task_type_insp), { size: 14, bold: true }),
        makeCell(chk(r.activity_training), { size: 14, bold: true }),
        makeCell(chk(r.activity_perform), { size: 14, bold: true }),
        makeCell(chk(r.activity_supervise), { size: 14, bold: true }),
        makeCell(chk(r.activity_crs), { size: 14, bold: true }),
        makeCell(t.ata || '', { size: 12 }),
        makeCell(t.operation_performed || '', { size: 12, align: AlignmentType.LEFT }),
        makeCell(t.time_duration || '', { size: 12 }),
        makeCell(t.maintenance_record_ref || '', { size: 11, align: AlignmentType.LEFT }),
        makeCell(r.remarks || '', { size: 11, align: AlignmentType.LEFT }),
      ]
    })
  })

  const mainTable = new Table({
    width: { size: PW, type: WidthType.DXA },
    columnWidths: allColWidths,
    rows: [hdr1, hdr2, ...taskRows]
  })

  const notePara = new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 60 },
    children: [
      new TextRun({ text: 'Note : ', bold: true, size: 11, italics: true, font: 'Arial' }),
      new TextRun({ text: 'The tasks recorded need to be related to a combination of activities / authorizations and not limited to simple tasks.  /  ', italics: true, size: 11, font: 'Arial' }),
      new TextRun({ text: 'Note : ', bold: true, size: 11, italics: true, font: 'Arial' }),
      new TextRun({ text: "Les tâches enregistrées doivent être associées à une combinaison d'activités / autorisations et ne se limitent pas à des tâches simples.", italics: true, size: 11, font: 'Arial' }),
    ]
  })

  const sigTable = new Table({
    width: { size: PW, type: WidthType.DXA },
    columnWidths: [Math.round(PW / 2), Math.round(PW / 2)],
    rows: [new TableRow({
      height: { value: 900, rule: HeightRule.ATLEAST },
      children: [
        new TableCell({
          borders: allBorders(),
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [
            new Paragraph({ children: [new TextRun({ text: "Logbook owner's signature (*) :", bold: true, size: 14, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: 'Signature du détenteur du livret', italics: true, size: 12, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: ' ', size: 10 })] }),
            new Paragraph({ children: [new TextRun({ text: 'Date :', size: 13, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: '…………………………….', size: 13, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: 'Name (Nom) :', size: 13, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: `${user?.name || ''} ${user?.surname || ''}`, size: 13, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: 'Signature :', size: 13, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: '…………………………….', size: 13, font: 'Arial' })] }),
          ]
        }),
        new TableCell({
          borders: allBorders(),
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [
            new Paragraph({ children: [new TextRun({ text: "Team Leader / Manager's validation :", bold: true, size: 14, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: "Validation du Chef d'Equipe / Responsable", italics: true, size: 12, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: ' ', size: 10 })] }),
            new Paragraph({ children: [new TextRun({ text: 'Date :', size: 13, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: '…………………………….', size: 13, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: 'Name (Nom) :', size: 13, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: '…………………………….', size: 13, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: 'Signature :', size: 13, font: 'Arial' })] }),
            new Paragraph({ children: [new TextRun({ text: '…………………………….', size: 13, font: 'Arial' })] }),
          ]
        })
      ]
    })]
  })

  const declarationPara = new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 60 },
    children: [
      new TextRun({ text: '(*) I declare that the entries in this logbook are complete and true. / Je ', italics: true, size: 11, font: 'Arial' }),
      new TextRun({ text: 'déclare', italics: true, size: 11, font: 'Arial', underline: { type: UnderlineType.SINGLE } }),
      new TextRun({ text: ' que les ', italics: true, size: 11, font: 'Arial' }),
      new TextRun({ text: 'enregistrements', italics: true, size: 11, font: 'Arial', underline: { type: UnderlineType.SINGLE } }),
      new TextRun({ text: ' de ce ', italics: true, size: 11, font: 'Arial' }),
      new TextRun({ text: 'livret', italics: true, size: 11, font: 'Arial', underline: { type: UnderlineType.SINGLE } }),
      new TextRun({ text: ' sont ', italics: true, size: 11, font: 'Arial' }),
      new TextRun({ text: 'complets', italics: true, size: 11, font: 'Arial', underline: { type: UnderlineType.SINGLE } }),
      new TextRun({ text: ' et ', italics: true, size: 11, font: 'Arial' }),
      new TextRun({ text: 'vrais', italics: true, size: 11, font: 'Arial', underline: { type: UnderlineType.SINGLE } }),
      new TextRun({ text: '.', italics: true, size: 11, font: 'Arial' }),
    ]
  })

  const footerPara = new Paragraph({
    spacing: { before: 40 },
    children: [
      new TextRun({ text: 'AHDT - 101', italics: true, size: 11, font: 'Arial' }),
      new TextRun({ text: '\t\t\t\t\t\t\t\t- Sept. 2018 -', italics: true, size: 11, font: 'Arial' }),
    ]
  })

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 16838, height: 11906, orientation: PageOrientation.LANDSCAPE },
          margin: { top: 500, right: 500, bottom: 500, left: 500 }
        }
      },
      children: [
        headerTable,
        new Paragraph({ spacing: { before: 60, after: 0 }, children: [new TextRun({ text: '' })] }),
        leftNote, recordsPara, mainTable,
        new Paragraph({ spacing: { before: 40, after: 0 }, children: [new TextRun({ text: '' })] }),
        notePara, sigTable, declarationPara, footerPara,
      ]
    }]
  })
}

// ── WORD ROUTE ────────────────────────────────────────────────────────────────
router.get('/entry/:entryId', async (req, res) => {
  try {
    const { entryId } = req.params
    const { data: entry } = await supabase
      .from('log_entries').select('*, protocols(*, aircraft_types(*))').eq('id', entryId).single()
    if (!entry) return res.status(404).json({ error: 'Not found' })

    const { data: tasks } = await supabase
      .from('tasks').select('*').eq('protocol_id', entry.protocol_id).neq('is_active', false).order('order')
    const { data: entryRows } = await supabase
      .from('log_entry_rows').select('*').eq('log_entry_id', entryId)
    const { data: user } = await supabase
      .from('users').select('*').eq('id', entry.user_id).single()

    const rowMap = {}
    entryRows?.forEach(r => { rowMap[r.task_id] = r })

    const doc = buildDoc(entry, entry.protocols, tasks || [], rowMap, user)
    const buffer = await Packer.toBuffer(doc)
    const filename = `carnet_${entry.protocols?.aircraft_types?.name}_${entry.date}.docx`
      .replace(/\s/g, '_').replace(/[^a-zA-Z0-9_.-]/g, '')

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`)
    res.send(buffer)
  } catch (err) {
    console.error(err)
    res.status(500).json({ error: err.message })
  }
})

// ── PDF ROUTE (single entry) ──────────────────────────────────────────────────
router.get('/entry/:entryId/pdf', async (req, res) => {
  try {
    const { entryId } = req.params
    const { data: entry } = await supabase
      .from('log_entries').select('*, protocols(*, aircraft_types(*))').eq('id', entryId).single()
    if (!entry) return res.status(404).json({ error: 'Not found' })

    const { data: tasks } = await supabase
      .from('tasks').select('*').eq('protocol_id', entry.protocol_id).neq('is_active', false).order('order')
    const { data: entryRows } = await supabase
      .from('log_entry_rows').select('*').eq('log_entry_id', entryId)
    const { data: user } = await supabase
      .from('users').select('*').eq('id', entry.user_id).single()

    const rowMap = {}
    entryRows?.forEach(r => { rowMap[r.task_id] = r })

    const aircraft = entry.ac_type || entry.protocols?.aircraft_types?.name || ''
    const rating = entry.rating || entry.protocols?.rating || ''
    const privilege = entry.privilege || entry.protocols?.privilege || ''
    const logoUrl = await getLogoUrl()
    const sigName = `${user?.name || ''} ${user?.surname || ''}`

    const taskRowsRaw = (tasks || []).map(t => {
      const r = rowMap[t.id] || {}
      return `<tr>
        <td>${entry.date || ''}</td><td>${entry.location || ''}</td>
        <td>${aircraft}</td><td>${entry.ac_registration || ''}</td>
        <td>${rating}</td><td>${privilege}</td>
        <td class="check">${chk(r.task_type_fot)}</td>
        <td class="check">${chk(r.task_type_sgh)}</td>
        <td class="check">${chk(r.task_type_ri)}</td>
        <td class="check">${chk(r.task_type_mel)}</td>
        <td class="check">${chk(r.task_type_ts)}</td>
        <td class="check">${chk(r.task_type_mod)}</td>
        <td class="check">${chk(r.task_type_rep)}</td>
        <td class="check">${chk(r.task_type_insp)}</td>
        <td class="check">${chk(r.activity_training)}</td>
        <td class="check">${chk(r.activity_perform)}</td>
        <td class="check">${chk(r.activity_supervise)}</td>
        <td class="check">${chk(r.activity_crs)}</td>
        <td>${t.ata || ''}</td>
        <td class="op-cell">${t.operation_performed || ''}</td>
        <td>${t.time_duration || ''}</td>
        <td class="ref-cell">${t.maintenance_record_ref || ''}</td>
        <td>${r.remarks || ''}</td>
      </tr>`
    })

    const taskRowsHtml = padToTen(taskRowsRaw.join(''), taskRowsRaw.length)

    const headerHtml = `
      <table class="doc-table header-table">
        <tr>
          <td class="header-logo">
            ${logoUrl ? `<img src="${logoUrl}" class="logo-img" alt="Air Algérie">` : '<div style="color:#CC0000;font-weight:bold;font-size:11px;">AIR ALGÉRIE</div>'}
            <div class="logo-technics">TECHNICS</div>
            <div class="logo-ref">DACM.AMO.01 / EASA.145.0163</div>
          </td>
          <td class="header-title">
            <div class="title-main">MAINTENANCE EXPERIENCE LOGBOOK</div>
            <div class="title-sub">LIVRET D'EXPERIENCE EN MAINTENANCE</div>
          </td>
          <td class="header-ids">
            <div><strong>AHT ID :</strong>&nbsp;${user?.aht_id || '………………'}</div>
            <div><strong>Page :</strong>&nbsp;………………</div>
          </td>
        </tr>
      </table>`

    const html = `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><style>${CSS}</style></head>
<body>${buildPageHtml(headerHtml, taskRowsHtml, sigName)}</body></html>`

    const puppeteer = await import('puppeteer')
    const browser = await puppeteer.default.launch({ args: ['--no-sandbox', '--disable-setuid-sandbox'] })
    const page = await browser.newPage()
    await page.setContent(html, { waitUntil: 'networkidle0' })
    const pdf = await page.pdf({
      format: 'A4', landscape: true, printBackground: true,
      margin: { top: '0', right: '0', bottom: '0', left: '0' }
    })
    await browser.close()

    const filename = `carnet_${entry.protocols?.aircraft_types?.name}_${entry.date}.pdf`
      .replace(/\s/g, '_').replace(/[^a-zA-Z0-9_.-]/g, '')

    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`)
    res.send(pdf)
  } catch (err) {
    console.error(err)
    res.status(500).json({ error: err.message })
  }
})

// ── FULL CARNET PDF ───────────────────────────────────────────────────────────
router.get('/carnet/:userId/pdf', async (req, res) => {
  try {
    const { userId } = req.params
    const { data: user } = await supabase.from('users').select('*').eq('id', userId).single()
    const { data: entries } = await supabase
      .from('log_entries').select('*, protocols(*, aircraft_types(*))')
      .eq('user_id', userId).order('date', { ascending: true })
    const { data: standaloneTasks } = await supabase
      .from('log_entry_rows').select('*')
      .eq('standalone_user_id', userId).eq('is_standalone', true)
      .order('standalone_date', { ascending: true })

    const logoUrl = await getLogoUrl()
    const sigName = `${user?.name || ''} ${user?.surname || ''}`

    const logoTag = logoUrl
      ? `<img src="${logoUrl}" class="logo-img" alt="Air Algérie">`
      : `<div style="color:#CC0000;font-weight:bold;font-size:11px;">AIR ALGÉRIE</div>`

    function headerHtml(pageNum) {
      return `
        <table class="doc-table header-table">
          <tr>
            <td class="header-logo">
              ${logoTag}
              <div class="logo-technics">TECHNICS</div>
              <div class="logo-ref">DACM.AMO.01 / EASA.145.0163</div>
            </td>
            <td class="header-title">
              <div class="title-main">MAINTENANCE EXPERIENCE LOGBOOK</div>
              <div class="title-sub">LIVRET D'EXPERIENCE EN MAINTENANCE</div>
            </td>
            <td class="header-ids">
              <div><strong>AHT ID :</strong> ${user?.aht_id || '………………'}</div>
              <div><strong>Page :</strong> ${pageNum}</div>
            </td>
          </tr>
        </table>`
    }

    const coverPage = `
      <div class="page">
        ${headerHtml(1)}
        <div class="personnel-title">
          <strong>PERSONNEL DATA</strong><br><em>INFORMATION SUR LA PERSONNE</em>
        </div>
        <table class="doc-table personnel-table">
          <tr><td class="pers-label"><strong>Name</strong><br><em>Nom</em></td><td>${user?.surname || ''}</td></tr>
          <tr><td class="pers-label"><strong><u>Surname</u></strong><br><em>Prénom</em></td><td>${user?.name || ''}</td></tr>
          <tr><td class="pers-label"><strong>Date &amp; place of birth</strong><br><em>Date et lieu de naissance</em></td>
            <td>${user?.date_of_birth || ''} ${user?.place_of_birth ? '— ' + user.place_of_birth : ''}</td></tr>
          <tr><td class="pers-label"><strong>Authorization Nr</strong><br><em>Numéro d'habilitation</em></td><td>${user?.authorization_nr || ''}</td></tr>
          <tr><td class="pers-label"><strong>Scope of autorization</strong><br><em>Domaine d'habilitation</em></td><td>${user?.scope_of_authorization || ''}</td></tr>
          <tr><td class="pers-label"><strong>Signature</strong><br><em>Signature</em></td><td style="height:40px"></td></tr>
        </table>
        <div class="footer-ref"><span>AHDT - 101</span><span>- Sept. 2018 -</span></div>
      </div>`

    let pageNum = 2
    const entryPages = []

    for (const entry of (entries || [])) {
      const { data: tasks } = await supabase
        .from('tasks').select('*').eq('protocol_id', entry.protocol_id).neq('is_active', false).order('order')
      const { data: entryRows } = await supabase
        .from('log_entry_rows').select('*').eq('log_entry_id', entry.id)

      const rowMap = {}
      entryRows?.forEach(r => { rowMap[r.task_id] = r })

      const aircraft = entry.ac_type || entry.protocols?.aircraft_types?.name || ''
      const rating = entry.rating || entry.protocols?.rating || ''
      const privilege = entry.privilege || entry.protocols?.privilege || ''

      const chunks = []
      for (let i = 0; i < Math.max((tasks || []).length, 1); i += TASKS_PER_PAGE) {
        chunks.push((tasks || []).slice(i, i + TASKS_PER_PAGE))
      }

      for (const chunk of chunks) {
        const rawRows = chunk.map(t => {
          const r = rowMap[t.id] || {}
          const c = v => v ? 'X' : ''
          return `<tr>
            <td>${entry.date || ''}</td><td>${entry.location || ''}</td>
            <td>${aircraft}</td><td>${entry.ac_registration || ''}</td>
            <td>${rating}</td><td>${privilege}</td>
            <td class="check">${c(r.task_type_fot)}</td><td class="check">${c(r.task_type_sgh)}</td>
            <td class="check">${c(r.task_type_ri)}</td><td class="check">${c(r.task_type_mel)}</td>
            <td class="check">${c(r.task_type_ts)}</td><td class="check">${c(r.task_type_mod)}</td>
            <td class="check">${c(r.task_type_rep)}</td><td class="check">${c(r.task_type_insp)}</td>
            <td class="check">${c(r.activity_training)}</td><td class="check">${c(r.activity_perform)}</td>
            <td class="check">${c(r.activity_supervise)}</td><td class="check">${c(r.activity_crs)}</td>
            <td>${t.ata || ''}</td>
            <td class="op-cell">${t.operation_performed || ''}</td>
            <td>${t.time_duration || ''}</td>
            <td class="ref-cell">${t.maintenance_record_ref || ''}</td>
            <td>${r.remarks || ''}</td>
          </tr>`
        })
        const bodyHtml = padToTen(rawRows.join(''), rawRows.length)
        entryPages.push(buildPageHtml(headerHtml(pageNum++), bodyHtml, sigName))
      }
    }

    // Standalone task pages
    const standaloneChunks = []
    for (let i = 0; i < (standaloneTasks || []).length; i += TASKS_PER_PAGE) {
      standaloneChunks.push((standaloneTasks || []).slice(i, i + TASKS_PER_PAGE))
    }
    for (const chunk of standaloneChunks) {
      const rawRows = chunk.map(r => {
        const c = v => v ? 'X' : ''
        return `<tr>
          <td>${r.standalone_date || ''}</td><td>${r.standalone_location || ''}</td>
          <td>${r.standalone_ac_type || ''}</td><td>${r.standalone_ac_registration || ''}</td>
          <td>${r.standalone_rating || ''}</td><td>${r.standalone_privilege || ''}</td>
          <td class="check">${c(r.task_type_fot)}</td><td class="check">${c(r.task_type_sgh)}</td>
          <td class="check">${c(r.task_type_ri)}</td><td class="check">${c(r.task_type_mel)}</td>
          <td class="check">${c(r.task_type_ts)}</td><td class="check">${c(r.task_type_mod)}</td>
          <td class="check">${c(r.task_type_rep)}</td><td class="check">${c(r.task_type_insp)}</td>
          <td class="check">${c(r.activity_training)}</td><td class="check">${c(r.activity_perform)}</td>
          <td class="check">${c(r.activity_supervise)}</td><td class="check">${c(r.activity_crs)}</td>
          <td>${r.standalone_ata || ''}</td>
          <td class="op-cell">${r.standalone_operation || ''}</td>
          <td>${r.time_duration || ''}</td>
          <td class="ref-cell">${r.standalone_ref || ''}</td>
          <td>${r.remarks || ''}</td>
        </tr>`
      })
      const bodyHtml = padToTen(rawRows.join(''), rawRows.length)
      entryPages.push(buildPageHtml(headerHtml(pageNum++), bodyHtml, sigName))
    }

    const html = `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><style>${CSS}</style></head>
<body>${coverPage}${entryPages.join('')}</body></html>`

    const puppeteer = await import('puppeteer')
    const browser = await puppeteer.default.launch({ args: ['--no-sandbox', '--disable-setuid-sandbox'] })
    const page = await browser.newPage()
    await page.setContent(html, { waitUntil: 'networkidle0' })
    const pdf = await page.pdf({
      format: 'A4', landscape: true, printBackground: true,
      margin: { top: '0', right: '0', bottom: '0', left: '0' }
    })
    await browser.close()

    const filename = `carnet_${user?.surname}_${user?.name}.pdf`
      .replace(/\s/g, '_').replace(/[^a-zA-Z0-9_.-]/g, '')

    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`)
    res.send(pdf)
  } catch (err) {
    console.error(err)
    res.status(500).json({ error: err.message })
  }
})

export default router