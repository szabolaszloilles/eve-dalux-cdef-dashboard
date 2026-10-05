"""Build a traceable CDEF deadline summary without changing the source workbook.

Requires openpyxl. Run: python build_summary.py
CREC shorthand defaults to September 20/30, as confirmed by the user.
Hidden source rows are excluded from the report.
The default report uses contractor-list deadlines only; Data is a reference.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import to_excel
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.properties import CalcProperties


NAVY = '17365D'
BLUE = 'DCEAF7'
RED = 'FCE4D6'
AMBER = 'FFF2CC'
GREEN = 'E2F0D9'
DATE_FORMAT = 'dd mmm yyyy'
FORMULA_CACHE = {}
HEADER_ALIASES = {
    'no.': 'no.',
    'deadline': 'deadline',
    'correction deadline': 'deadline',
    'subject': 'subject',
    'problem description': 'problem description',
    'person in charge': 'person in charge',
    'responsible': 'responsible',
    'has taken action': 'has taken action',
    'status': 'status',
    'status detailed': 'status detailed',
    'work package': 'work package',
}


def norm(value):
    return str(value or '').strip().casefold()


def header_key(value):
    """Recognize known full labels within slash-separated bilingual headers."""
    label = norm(value)
    if not isinstance(value, str):
        return label
    for part in [label, *label.split('/')]:
        key = ' '.join(part.split())
        if key in HEADER_ALIASES:
            return HEADER_ALIASES[key]
    return label


def raw_text(value):
    if isinstance(value, datetime):
        return value.isoformat(sep=' ')
    return str(value) if value is not None else ''


def parse_deadline(value, year, crec_days):
    """Only accept observed, unambiguous formats; retain undecidable values."""
    if isinstance(value, datetime):
        return value.date(), ''
    if isinstance(value, date):
        return value, ''
    text = str(value).strip()
    try:
        match = re.fullmatch(r'(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})', text)
        if match:
            return date(*map(int, match.groups())), ''
        match = re.fullmatch(r'(\d{1,2})/(\d{1,2})/(\d{4})', text)
        if match:
            month, day, full_year = map(int, match.groups())
            if day <= 12 and month <= 12 and day != month:
                return None, 'Ambiguous day/month order; confirm date.'
            if month > 12:
                month, day = day, month
            return date(full_year, month, day), ''
        match = re.fullmatch(r'(\d{1,2})(\.+)(\d{1,2})', text)
        if match:
            month, separator, day_text = match.groups()
            month, day = int(month), int(day_text)
            note = f'Year assumed {year} from workbook context.'
            if isinstance(value, (float, int)) and len(day_text) == 1 and day in (1, 2, 3):
                if crec_days == 'review':
                    return None, f'Numeric {text}: day could be {day} or {day * 10}; confirm. Excel displays two decimal places.'
                day = day * 10 if crec_days == 'tens' else day
                note += ' Numeric shorthand interpretation confirmed by user.'
            if len(separator) > 1:
                note += f' Repeated dot in {text} normalized.'
            return date(year, month, day), note
    except ValueError:
        return None, 'Invalid calendar date; confirm original entry.'
    explanations = {
        'finish': 'Completion note, no date. Confirm completion; source export says Ongoing.',
        '\u672a\u627e\u5230': 'Text means "not found"; no date supplied.',
        '\u5df2\u56de\u590d': 'Text means "already replied"; no date supplied.',
    }
    return None, explanations.get(text.casefold(), 'Text in Deadline field; no calendar date supplied.')


def read_source(path, year, crec_days):
    workbook = load_workbook(path, data_only=True)
    entries, inventory = [], []
    for sheet in workbook:
        candidates = []
        for row in sheet.iter_rows():
            for cell in row:
                if header_key(cell.value) == 'deadline':
                    mapping = {header_key(c.value): c.column for c in sheet[cell.row] if c.value is not None}
                    if 'no.' in mapping:
                        candidates.append((cell, mapping))
        if len(candidates) > 1:
            raise ValueError(f'Multiple Deadline headers in {sheet.title}; review layout.')
        if not candidates:
            inventory.append([sheet.title, '', 0, 0, 0, 'Excluded: no Deadline header'])
            continue
        header, mapping = candidates[0]
        sheet_entries = []
        for row_number in range(header.row + 1, sheet.max_row + 1):
            def field(name):
                column = mapping.get(name)
                return sheet.cell(row_number, column).value if column else None

            cdef = str(field('no.') or '').strip()
            if not cdef:
                continue
            row_dimension = sheet.row_dimensions.get(row_number)
            hidden = bool(row_dimension and row_dimension.hidden)
            raw = sheet.cell(row_number, header.column).value
            filled = raw is not None and str(raw).strip() != ''
            deadline, issue = parse_deadline(raw, year, crec_days) if filled and not hidden else (None, '')
            owner_field = next((key for key in ['person in charge', 'responsible', 'has taken action'] if field(key)), '')
            entry = {
                'id': cdef, 'sheet': sheet.title, 'row': row_number, 'hidden': hidden,
                'cell': f'{header.column_letter}{row_number}', 'header': header.coordinate,
                'raw': raw, 'filled': filled, 'date': deadline, 'issue': issue,
                'subject': field('subject') or field('problem description') or '',
                'owner': field(owner_field) if owner_field else '', 'owner_field': owner_field,
                'status': field('status') or '', 'status_detailed': field('status detailed') or '',
                'work_package': field('work package') or '',
            }
            entries.append(entry)
            sheet_entries.append(entry)
        visible_entries = [e for e in sheet_entries if not e['hidden']]
        hidden_count = len(sheet_entries) - len(visible_entries)
        treatment = 'Included as source' if any(e['filled'] for e in visible_entries) else 'Excluded: no visible deadline values'
        if hidden_count:
            treatment += f'; {hidden_count} hidden CDEF rows excluded'
        inventory.append([
            sheet.title, header.coordinate, len(visible_entries),
            sum(e['filled'] for e in visible_entries), sum(e['date'] is not None for e in visible_entries),
            treatment,
        ])
    workbook.close()
    return entries, inventory


def choose_records(entries, policy):
    by_id = defaultdict(list)
    for entry in entries:
        by_id[entry['id']].append(entry)
    selected = []
    for cdef, all_rows in by_id.items():
        # Membership and deadlines come from visible contractor lists only.
        # Data may enrich an existing CDEF, but must not introduce one.
        contractors = [e for e in all_rows if e['sheet'] != 'Data']
        visible_rows = [e for e in all_rows if not e['hidden']]
        candidates = [e for e in contractors if not e['hidden'] and e['filled']]
        chosen = [e for e in visible_rows if e['filled']] if policy == 'all' else candidates[:1]
        for entry in chosen:
            others = [e for e in visible_rows if e is not entry]
            entry['company'] = 'Synergy' if entry['sheet'] == 'Synergy scope' else entry['sheet']
            if entry['sheet'] == 'Data':
                entry['company'] = str(entry['work_package'] or 'Data - unspecified')
            for key in ['subject', 'owner', 'status']:
                source = entry if entry[key] else next((e for e in others if e[key]), None)
                entry[f'display_{key}'] = source[key] if source else 'Not supplied'
                entry[f'{key}_source'] = f"{source['sheet']} row {source['row']}" if source else ''
            entry['other_deadlines'] = '; '.join(
                f"{e['sheet']}!{e['cell']}: {raw_text(e['raw'])}" for e in others if e['filled']
            )
            differing = [e for e in others if e['date'] and entry['date'] and e['date'] != entry['date']]
            entry['selection_note'] = ''
            if differing:
                entry['selection_note'] = 'Different source deadline retained in Source deadlines.'
            other_contractors = sorted({
                e['sheet'] for e in others
                if e['filled'] and e['sheet'] != 'Data'
                and ('Synergy' if e['sheet'] == 'Synergy scope' else e['sheet']) != entry['company']
            })
            entry['assignment_review'] = bool(other_contractors)
            if other_contractors:
                entry['selection_note'] += f" Also listed in {', '.join(other_contractors)}; confirm contractor assignment."
            if entry['sheet'] == 'Data':
                entry['selection_note'] = 'Data source entry (explicit all policy).'
            selected.append(entry)
    return selected


def base_sheet(workbook, title, heading, subtitle, widths):
    sheet = workbook.create_sheet(title)
    sheet.sheet_view.showGridLines = False
    sheet.sheet_view.zoomScale = 85
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.sheet_properties.outlinePr.summaryRight = False
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(widths))
    sheet.cell(1, 1, heading).font = Font(name='Calibri', size=20, bold=True, color='FFFFFF')
    for cell in sheet[1]:
        cell.fill = PatternFill('solid', fgColor=NAVY)
    sheet.row_dimensions[1].height = 35
    sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(widths))
    sheet.cell(2, 1, subtitle).font = Font(name='Calibri', size=11, color='526477')
    sheet.cell(2, 1).alignment = Alignment(wrap_text=True, vertical='center')
    sheet.row_dimensions[2].height = 32
    for i, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(i)].width = width
    sheet.freeze_panes = 'C6'
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = 'landscape'
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = '1:5'
    sheet.oddFooter.center.text = 'Page &P of &N'
    return sheet


def table(sheet, headers, rows, name, start=5):
    for row_idx, values in enumerate([headers, *rows], start):
        for col_idx, value in enumerate(values, 1):
            cell = sheet.cell(row_idx, col_idx, value)
            if isinstance(value, str):
                cell.data_type = 's'
            cell.font = Font(name='Calibri', size=11)
            cell.alignment = Alignment(vertical='top', wrap_text=True)
            if isinstance(value, (date, datetime)):
                cell.number_format = DATE_FORMAT
            if row_idx == start:
                cell.fill = PatternFill('solid', fgColor=NAVY)
                cell.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
            elif row_idx % 2 == 0:
                cell.fill = PatternFill('solid', fgColor='F2F6FA')
        sheet.row_dimensions[row_idx].height = 32 if row_idx == start else 34
    if rows:
        end = start + len(rows)
        tab = Table(displayName=name, ref=f'A{start}:{get_column_letter(len(headers))}{end}')
        tab.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        sheet.add_table(tab)


def formula(sheet, cell, expression, cached):
    sheet[cell] = expression
    FORMULA_CACHE[(sheet.title, cell)] = cached


def link_source(sheet, coordinate, source, entry):
    cell = sheet[coordinate]
    cell.hyperlink = Hyperlink(ref=coordinate, target=source.name,
                               location=f"'{entry['sheet'].replace(chr(39), chr(39)*2)}'!{entry['cell']}")
    cell.font = Font(name='Calibri', size=11, color='0563C1', underline='single')


def cache_formula_results(path, workbook):
    """Cache calculated snapshot values for Excel previews; Excel recalculates on open."""
    ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    names = {f'xl/worksheets/sheet{i}.xml': s.title for i, s in enumerate(workbook, 1)}
    with ZipFile(path) as archive:
        files = [(item, archive.read(item.filename)) for item in archive.infolist()]
    with ZipFile(path, 'w', ZIP_DEFLATED) as archive:
        for item, data in files:
            sheet_name = names.get(item.filename)
            if sheet_name:
                root = ET.fromstring(data)
                changed = False
                for cell in root.findall('.//m:sheetData/m:row/m:c', ns):
                    key = (sheet_name, cell.attrib['r'])
                    if key not in FORMULA_CACHE:
                        continue
                    result = FORMULA_CACHE[key]
                    value = cell.find('m:v', ns)
                    if value is None:
                        value = ET.SubElement(cell, f"{{{ns['m']}}}v")
                    cell.set('t', 'str' if isinstance(result, str) else 'n')
                    value.text = str(result)
                    changed = True
                if changed:
                    data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            archive.writestr(item, data)


def bucket(entry, as_of):
    if norm(entry['display_status']) in {'closed', 'completed', 'resolved', 'finished'}:
        return 'Closed'
    remaining = (entry['date'] - as_of).days
    return 'Overdue' if remaining < 0 else 'Due today' if remaining == 0 else 'Due within 7 days' if remaining <= 7 else 'Due later'


def build(args):
    FORMULA_CACHE.clear()
    source, output = Path(args.source), Path(args.output)
    if source.resolve() == output.resolve():
        raise ValueError('Output must be different from the source workbook.')
    initial_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    as_of = date.fromisoformat(args.as_of) if args.as_of else date.today()
    entries, inventory = read_source(source, args.year, args.crec_days)
    hidden_count = sum(e['hidden'] for e in entries)
    selected = choose_records(entries, args.policy)
    if args.policy == 'contractor':
        for row in inventory:
            if row[0] == 'Data':
                row[5] = 'Reference only: excluded from summary deadlines'
    dated = sorted([e for e in selected if e['date']], key=lambda e: (e['date'], e['company'], e['id']))
    review = sorted([e for e in selected if not e['date']], key=lambda e: (e['company'], e['id']))
    companies = sorted({e['company'] for e in selected})
    assignment_notes = [
        f"{e['id']}: selected {e['sheet']}!{e['cell']} ({raw_text(e['raw'])}); other source deadlines: {e['other_deadlines']}."
        for e in selected if e['assignment_review']
    ]
    workbook = Workbook()
    workbook.remove(workbook.active)
    workbook.calculation = CalcProperties(calcId=191029, fullCalcOnLoad=True, forceFullCalc=True)
    workbook.properties.title = 'CDEF deadline follow-up'
    workbook.properties.subject = 'Deadline status by contractor with traceable source references'
    workbook.properties.creator = 'CDEF reporting'
    unit = 'unique CDEFs' if args.policy == 'contractor' else 'source entries (duplicates included)'

    summary = base_sheet(workbook, 'Summary', 'CDEF deadline follow-up',
                         f'{source.name} | Visible rows only; counts are {unit}. See Review for entries without a usable date.',
                         [24, 17, 17, 17, 20, 17, 17, 18, 19, 19])
    summary['A3'] = 'As of date'
    summary['B3'] = as_of
    summary['B3'].number_format = DATE_FORMAT
    summary['B3'].fill = PatternFill('solid', fgColor=BLUE)
    summary['B3'].comment = Comment('Change this date to recalculate deadline status. Source records refresh only when build_summary.py is rerun.', 'CDEF reporting')
    summary.merge_cells('C3:J3')
    summary['C3'] = 'Edit the blue date to update deadline status. Source status is the recorded export status.'
    summary['C3'].alignment = Alignment(wrap_text=True, vertical='center')
    summary.row_dimensions[3].height = 30

    detail = base_sheet(workbook, 'CDEF Follow-up', 'CDEFs with deadlines',
                        'Sorted by deadline. Filter contractor or deadline status. Negative days remaining means overdue.',
                        [22, 19, 52, 17, 19, 24, 17, 17, 32, 23, 18, 35, 45, 26])
    detail_headers = ['Contractor', 'CDEF No.', 'Subject / description', 'Deadline', 'Source status', 'Deadline status',
                      'Days remaining', 'Days overdue', 'Owner / source contact', 'Deadline worksheet', 'Source cell',
                      'Original deadline', 'Date / selection notes', 'Status source']
    detail_rows = [[e['company'], e['id'], e['display_subject'], e['date'], e['display_status'], '', '', '',
                    e['display_owner'], e['sheet'], e['cell'], raw_text(e['raw']),
                    ' '.join(x for x in [e['issue'], e['selection_note']] if x), e['status_source']] for e in dated]
    table(detail, detail_headers, detail_rows, 'CDEFDeadlineRegister')
    last = 5 + len(dated)
    for row, entry in enumerate(dated, 6):
        status = bucket(entry, as_of)
        days = (entry['date'] - as_of).days
        expression = (f'=IF(OR(E{row}="Closed",E{row}="Completed",E{row}="Resolved",E{row}="Finished"),"Closed",'
                      f'IF(D{row}<Summary!$B$3,"Overdue",IF(D{row}=Summary!$B$3,"Due today",'
                      f'IF(D{row}<=Summary!$B$3+7,"Due within 7 days","Due later"))))')
        formula(detail, f'F{row}', expression, status)
        formula(detail, f'G{row}', f'=D{row}-Summary!$B$3', days)
        formula(detail, f'H{row}', f'=IF(F{row}="Closed",0,MAX(0,Summary!$B$3-D{row}))', 0 if status == 'Closed' else max(0, -days))
        link_source(detail, f'K{row}', source, entry)
        detail[f'I{row}'].comment = Comment(f"Source: {entry['owner_source']}. Source field: {entry['owner_field'] or 'enriched from matching CDEF'}. Has taken action is retained as a source contact field, not treated as status.", 'CDEF reporting')
    if dated:
        for status, color in [('Overdue', RED), ('Due today', AMBER), ('Due within 7 days', AMBER), ('Due later', GREEN)]:
            detail.conditional_formatting.add(f'F6:F{last}', CellIsRule(operator='equal', formula=[f'"{status}"'], fill=PatternFill('solid', fgColor=color)))
    detail.column_dimensions.group('L', 'N', hidden=True)

    summary_headers = ['Contractor', 'Dated CDEFs', 'Overdue', 'Due today', 'Due within 7 days', 'Due later', 'Closed', 'Review entries', 'Earliest deadline', 'Latest deadline']
    summary_rows = []
    for company in companies:
        dates = [e['date'] for e in dated if e['company'] == company]
        summary_rows.append([company, 0, 0, 0, 0, 0, 0, sum(e['company'] == company for e in review), min(dates) if dates else '', max(dates) if dates else ''])
    table(summary, summary_headers, summary_rows, 'CDEFContractorSummary')
    for row, company in enumerate(companies, 6):
        records = [e for e in dated if e['company'] == company]
        counts = Counter(bucket(e, as_of) for e in records)
        formula(summary, f'B{row}', f'=COUNTIF(\'CDEF Follow-up\'!$A$6:$A${last},A{row})', len(records))
        for col, status in [('C', 'Overdue'), ('D', 'Due today'), ('E', 'Due within 7 days'), ('F', 'Due later'), ('G', 'Closed')]:
            formula(summary, f'{col}{row}', f'=COUNTIFS(\'CDEF Follow-up\'!$A$6:$A${last},A{row},\'CDEF Follow-up\'!$F$6:$F${last},"{status}")', counts[status])
    total_row = 6 + len(companies)
    summary.cell(total_row, 1, 'TOTAL')
    for col in range(2, 9):
        letter = get_column_letter(col)
        values = [FORMULA_CACHE.get(('Summary', f'{letter}{r}'), summary.cell(r, col).value) for r in range(6, total_row)]
        expression = f'=SUM({letter}6:{letter}{total_row-1})' if companies else '=0'
        formula(summary, f'{letter}{total_row}', expression, sum(values))
    for cell in summary[total_row]:
        cell.fill = PatternFill('solid', fgColor=BLUE)
        cell.font = Font(name='Calibri', bold=True, size=11)
    summary.row_dimensions[total_row].height = 28
    summary.merge_cells(start_row=total_row + 2, start_column=1, end_row=total_row + 2, end_column=10)
    summary.cell(total_row + 2, 1, f'{len(dated)} dated {unit}; {len(review)} entries require date review. {hidden_count} hidden CDEF source rows excluded. See Notes for selection rules.').alignment = Alignment(wrap_text=True)
    summary.row_dimensions[total_row + 2].height = 32
    chart = BarChart()
    chart.type = 'bar'
    chart.grouping = 'stacked'
    chart.overlap = 100
    chart.title = 'Deadline status by contractor'
    if companies:
        chart.add_data(Reference(summary, min_col=3, max_col=7, min_row=5, max_row=total_row - 1), titles_from_data=True)
        chart.set_categories(Reference(summary, min_col=1, min_row=6, max_row=total_row - 1))
    chart.height, chart.width = 12, 27
    chart.x_axis.title = 'CDEF count'
    chart.legend.position = 'b'
    for series, color in zip(chart.series, ['C0504D', 'F4B183', 'FFD966', '70AD47', 'A5A5A5']):
        series.graphicalProperties.solidFill = color
    if companies:
        summary.add_chart(chart, f'A{total_row + 4}')
    summary.print_area = f'A1:J{total_row + 29}'

    schedule = base_sheet(workbook, 'By deadline', 'Deadline schedule', 'Number of dated CDEFs per deadline and contractor; the Review sheet contains excluded non-date entries.', [18, *([18] * len(companies)), 17])
    date_counts = Counter((e['date'], e['company']) for e in dated)
    date_rows = []
    for deadline in sorted({e['date'] for e in dated}):
        counts = [date_counts[deadline, company] for company in companies]
        date_rows.append([deadline, *counts, sum(counts)])
    table(schedule, ['Deadline', *companies, 'Total'], date_rows, 'CDEFByDeadline')

    review_sheet = base_sheet(workbook, 'Review', 'Deadline entries to review', 'These entries are excluded from dated counts. Resolve the original Deadline value, then rebuild the summary.', [23, 19, 48, 28, 62, 20, 22, 18])
    review_rows = [[e['company'], e['id'], e['display_subject'], raw_text(e['raw']), e['issue'], e['display_status'], e['sheet'], e['cell']] for e in review]
    table(review_sheet, ['Contractor', 'CDEF No.', 'Subject / description', 'Original deadline', 'Review reason', 'Source status', 'Worksheet', 'Source cell'], review_rows, 'CDEFDeadlineReview')
    for row, entry in enumerate(review, 6):
        review_sheet.row_dimensions[row].height = 45
        review_sheet.cell(row, 4).fill = PatternFill('solid', fgColor=AMBER)
        link_source(review_sheet, f'H{row}', source, entry)

    audit = base_sheet(workbook, 'Source deadlines', 'Populated Deadline fields from visible rows', 'Audit records include visible source duplicates and are not the unique CDEF total. Hidden source rows are excluded.', [23, 18, 19, 30, 18, 30, 60, 20])
    picked = {id(e) for e in selected}
    selected_ids = {e['id'] for e in selected}
    audit_entries = [e for e in entries if not e['hidden'] and e['filled']]
    audit_rows = []
    for entry in audit_entries:
        use = 'Dated register' if id(entry) in picked and entry['date'] else 'Review' if id(entry) in picked else 'Alternate source' if entry['id'] in selected_ids else 'Excluded: Data is reference only' if entry['sheet'] == 'Data' and args.policy == 'contractor' else 'Excluded: no visible contractor deadline'
        audit_rows.append([entry['sheet'], entry['cell'], entry['id'], raw_text(entry['raw']), entry['date'], use, entry['issue'], entry['status'] or 'Not supplied'])
    table(audit, ['Worksheet', 'Source cell', 'CDEF No.', 'Original deadline', 'Parsed deadline', 'Use in summary', 'Date interpretation', 'Source status'], audit_rows, 'CDEFSourceDeadlines')
    for row, entry in enumerate(audit_entries, 6):
        link_source(audit, f'B{row}', source, entry)

    notes = base_sheet(workbook, 'Notes', 'How to use this workbook', 'Source selection, date assumptions and refresh instructions.', [27, 105, 20, 20, 20, 36])
    explanations = [
        ('Purpose', 'Summary and By deadline summarize dated CDEFs. CDEF Follow-up is the filterable working register. Review lists non-date or ambiguous entries. Source deadlines preserves populated Deadline cells from visible rows.'),
        ('Counting rule', 'One record per CDEF No. from visible contractor lists in Summary and CDEF Follow-up. Data-only CDEFs are excluded.' if args.policy == 'contractor' else 'Every populated visible worksheet Deadline entry is included; totals count source records and repeat CDEF IDs.'),
        ('Hidden rows', f'{hidden_count} hidden CDEF source rows are excluded from deadline selection, status/contact enrichment, Review and Source deadlines. This includes rows hidden manually, by grouping or by a saved Excel filter. Visible copies on other contractor sheets remain eligible.'),
        ('Deadline selection', 'Recognize Deadline or Correction Deadline at any column/header row, including slash-separated bilingual labels. No. and Subject also support bilingual labels. Deadline2, Construction Deadline and Planned Time are not substituted. Multiple recognized deadline columns require layout review.'),
        ('Duplicate selection', 'Use only visible populated contractor Deadline entries. Synergy scope precedes Synergy where both provide a visible deadline. Data never supplies a summary deadline or adds a CDEF. It can supply missing subject/contact/status details for an existing selected CDEF. Visible reference entries remain in Source deadlines.' if args.policy == 'contractor' else 'No source precedence is applied. Duplicate CDEF IDs remain separate visible source records, including Data; hidden source rows are excluded.'),
        ('Conflicting dates', 'Where visible sources disagree, the main summary uses the selected contractor deadline under the contractor policy. Synergy scope takes priority over Synergy. Alternate visible source dates are retained in Source deadlines.'),
        ('Contractor assignment', ' '.join(assignment_notes) if assignment_notes else 'No selected CDEF is listed with a deadline under different contractors.'),
        ('No deadlines', 'Blank Deadline rows are excluded. Minimax has no populated Deadline values and is excluded. A blank contractor deadline does not fall back to an old Data deadline.'),
        ('Date formats', f'Real Excel dates and explicit year/month/day text are converted to calendar dates. Times are retained in Original deadline, while status uses calendar days. Yearless CREC dates assume {args.year}, consistent with the August 2026 source export.'),
        ('CREC shorthand', {'review': 'Numeric 9.2 / 9.3 entries are held in Review because they could mean September 2/3 or 20/30. Excel displays them as 9.20 / 9.30. Confirm before assigning a date.', 'tens': 'User confirmed numeric 9.2 / 9.3 mean September 20 / 30.', 'single': 'User confirmed numeric 9.2 / 9.3 mean September 2 / 3.'}[args.crec_days]),
        ('CREC typo', 'Ten entries written as 9..15 are interpreted as September 15 with the assumed year. The normalization is recorded against each affected CDEF.'),
        ('Non-date notes', 'finish, Chinese reply/not-found notes and an out-of-scope note are preserved in Review, not converted into dates or inferred completion statuses.'),
        ('Status meaning', 'Source status is copied from the selected visible row or a matching visible CDEF in another sheet; the source is shown. The original export says 31 Aug 2026 and all supplied Status values are Ongoing. Deadline status indicates timing, not verified completion.'),
        ('Contact fields', 'Owner / source contact uses Person in charge, Responsible or Has taken action as available. The last field contains names in CSCEC5 INFRA. Each contact cell has a comment giving its source.'),
        ('Update date', 'Change Summary!B3 to recalculate overdue, due today, within 7 days and later counts. Due within 7 days means after the as-of date through as-of date + 7 days. Closed/Completed/Resolved/Finished statuses are excluded from overdue counts.'),
        ('Refresh source', 'After changing the original workbook, rerun python build_summary.py. This rebuilds the output from the source. The source workbook is never modified. To resolve CREC shorthand, add --crec-days tens or --crec-days single.'),
        ('Source links', 'Keep this summary beside CDEF Deadline.xlsx. Source cell links open the referenced worksheet and Deadline cell. Expand hidden columns L:N in CDEF Follow-up for original values, notes and status provenance.'),
        ('Source SHA-256', initial_hash),
        ('Build date', as_of.isoformat()),
    ]
    for row, (label, explanation) in enumerate(explanations, 5):
        notes.cell(row, 1, label).font = Font(name='Calibri', bold=True, color=NAVY)
        notes.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
        notes.cell(row, 2, explanation).alignment = Alignment(wrap_text=True, vertical='center')
        notes.row_dimensions[row].height = 42 if len(explanation) > 200 else 32
    inventory_row = 7 + len(explanations)
    table(notes, ['Source worksheet', 'Deadline header', 'Visible CDEF rows', 'Visible Deadline values', 'Visible parsed dates', 'Treatment'], inventory, 'CDEFSourceInventory', start=inventory_row)
    notes.freeze_panes = 'B5'
    workbook.active = 0
    workbook.save(output)
    cache_formula_results(output, workbook)
    assert hashlib.sha256(source.read_bytes()).hexdigest() == initial_hash, 'Source workbook changed during build.'
    result = {
        'output': str(output.resolve()), 'dated_count': len(dated), 'review_count': len(review),
        'source_deadline_entries': len(audit_entries), 'dated_by_contractor': dict(Counter(e['company'] for e in dated)),
        'hidden_source_rows_excluded': hidden_count,
        'deadline_status': dict(Counter(bucket(e, as_of) for e in dated)), 'as_of': str(as_of),
        'source_sha256': initial_hash, 'policy': args.policy, 'crec_days': args.crec_days,
    }
    print(json.dumps(result, ensure_ascii=True, indent=2))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', default='CDEF Deadline.xlsx')
    parser.add_argument('--output', default='CDEF Status Summary.xlsx')
    parser.add_argument('--as-of', help='As-of date in YYYY-MM-DD format (default: today).')
    parser.add_argument('--year', type=int, default=2026, help='Year assumed for yearless deadline values.')
    parser.add_argument('--crec-days', choices=['review', 'tens', 'single'], default='tens')
    parser.add_argument('--policy', choices=['contractor', 'all'], default='contractor',
                        help='contractor: visible contractor deadlines only (default); all: every visible source deadline, including Data and duplicates.')
    build(parser.parse_args())
