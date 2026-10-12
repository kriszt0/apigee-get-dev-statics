#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""Apigee period-based Excel exporter, Python 2.7 standard library only.
Uses installed psql and writes a real XLSX without third-party modules.
"""
from __future__ import print_function
import argparse
import csv
import datetime
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

DEFAULT_HOST = '/opt/apigee/var/run/apigee-postgresql'
HEADERS = ['api_product', 'apiproxy', 'developer_app',
           'message_count', 'error_count', 'error_rate_percent']
MAX_EXCEL_ROWS = 1048576

def xml_text(value):
    if not isinstance(value, unicode):
        value = value.decode('utf-8', 'replace')
    # XML 1.0 valid characters only
    value = u''.join(c for c in value if ord(c) in (9, 10, 13)
                     or 32 <= ord(c) <= 0xD7FF
                     or 0xE000 <= ord(c) <= 0xFFFD)
    return escape(value, {'"': '&quot;'})

def column_name(index):
    result = ''
    while index:
        index, part = divmod(index - 1, 26)
        result = chr(65 + part) + result
    return result

def cell(ref, data, kind='str', style=0):
    attr = ' s="%s"' % style if style else ''
    if data is None or data == '':
        return '<c r="%s"%s/>' % (ref, attr)
    if kind == 'num':
        try:
            # Numeric check: no untrusted content written as number
            float(data)
            return '<c r="%s"%s><v>%s</v></c>' % (ref, attr, data)
        except (ValueError, TypeError):
            pass
    return '<c r="%s" t="inlineStr"%s><is><t>%s</t></is></c>' % (
        ref, attr, xml_text(data).encode('utf-8'))

def allowed_table(name):
    # Accept schema."org.env.fact" or schema.table; nothing executable
    pattern = r'^[A-Za-z_][A-Za-z_0-9]*\.(?:"[A-Za-z_0-9.\-]+"|[A-Za-z_][A-Za-z_0-9]*)$'
    if not re.match(pattern, name):
        raise ValueError('Invalid --table format; example: analytics."ORG.ENV.fact"')
    return name

def query_for_period(table, period_start, period_end, label):
    # Exactly one result for each product + proxy + app within the period.
    # is_error is aggregated, not included in GROUP BY.
    return ("COPY ("
            "SELECT api_product, apiproxy, developer_app, "
            "SUM(COALESCE(message_count,0)) AS message_count, "
            "SUM(CASE WHEN is_error = 1 THEN COALESCE(message_count,0) ELSE 0 END) AS error_count, "
            "ROUND(100.0 * SUM(CASE WHEN is_error = 1 THEN COALESCE(message_count,0) ELSE 0 END) "
            "/ NULLIF(SUM(COALESCE(message_count,0)),0), 2) AS error_rate_percent "
            "FROM %s WHERE client_received_start_timestamp >= TIMESTAMP '%s' "
            "AND client_received_start_timestamp < TIMESTAMP '%s' "
            "GROUP BY api_product, apiproxy, developer_app "
            "ORDER BY message_count DESC, api_product, apiproxy, developer_app"
            ") TO STDOUT WITH (FORMAT CSV, HEADER TRUE)") % (
                table, period_start.strftime('%Y-%m-%d %H:%M:%S'),
                period_end.strftime('%Y-%m-%d %H:%M:%S'))

def add_month(value):
    year = value.year + (1 if value.month == 12 else 0)
    month = 1 if value.month == 12 else value.month + 1
    return value.replace(year=year, month=month, day=1)

def periods(start_date, end_date, granularity):
    """Inclusive input dates, buckets clipped to chosen date range."""
    first = datetime.datetime.combine(start_date, datetime.time.min)
    stop = datetime.datetime.combine(end_date + datetime.timedelta(days=1), datetime.time.min)
    current = first
    while current < stop:
        if granularity == 'hourly':
            natural = current + datetime.timedelta(hours=1)
            label = current.strftime('%Y-%m-%d %H')
        elif granularity == 'daily':
            natural = current + datetime.timedelta(days=1)
            label = current.strftime('%Y-%m-%d')
        elif granularity == 'weekly':
            monday = current - datetime.timedelta(days=current.weekday())
            natural = monday + datetime.timedelta(days=7)
            label = monday.strftime('%Y-%m-%d')
        else:
            natural = add_month(current.replace(day=1))
            label = current.strftime('%Y-%m')
        next_boundary = min(natural, stop)
        if next_boundary <= current:
            raise RuntimeError('Period calculation failed')
        yield current, next_boundary, label
        current = next_boundary

def write_sheet(z, sheet_no, period_start, period_end, label, opts):
    name = 'xl/worksheets/sheet%d.xml' % sheet_no
    sql = query_for_period(opts.table, period_start, period_end, label)
    cmd = ['psql', '-X', '-q', '-v', 'ON_ERROR_STOP=1',
           '-h', opts.host, '-U', opts.user, '-d', opts.database, '-c', sql]
    # Stream via temp file so psql errors cannot silently produce a broken workbook.
    tmp = tempfile.TemporaryFile()
    err = tempfile.TemporaryFile()
    proc = subprocess.Popen(cmd, stdout=tmp, stderr=err)
    code = proc.wait()
    if code != 0:
        err.seek(0)
        detail = err.read().decode('utf-8', 'replace')
        tmp.close(); err.close()
        raise RuntimeError('psql failed for %s (exit %d): %s' % (label, code, detail))
    err.close()
    tmp.seek(0)
    reader = csv.reader(tmp)
    try:
        source_headers = next(reader)
    except StopIteration:
        tmp.close()
        raise RuntimeError('Empty psql output for %s' % label)
    expected = HEADERS
    if source_headers != expected:
        tmp.close()
        raise RuntimeError('Unexpected columns for %s: %r' % (label, source_headers))
    count = 0
    sheet_tmp = tempfile.TemporaryFile()
    output = sheet_tmp
    try:
        output.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        output.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        output.write('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
        output.write('<cols><col min="1" max="3" width="28" customWidth="1"/>')
        output.write('')
        output.write('<col min="4" max="6" width="21" customWidth="1"/></cols>')
        output.write('<sheetData><row r="1">')
        for i, heading in enumerate(HEADERS, 1):
            output.write(cell('%s1' % column_name(i), heading, style=1))
        output.write('</row>')
        for row in reader:
            if len(row) != 6:
                raise RuntimeError('Unexpected CSV field count on %s: %d' % (label, len(row)))
            count += 1
            if count + 1 > MAX_EXCEL_ROWS:
                raise RuntimeError('Excel sheet exceeds 1048576 rows on %s' % label)
            r = count + 1
            output.write('<row r="%d">' % r)
            for i, data in enumerate(row, 1):
                kind = 'num' if i in (4, 5, 6) else 'str'
                output.write(cell('%s%d' % (column_name(i), r), data, kind=kind))
            output.write('</row>')
        output.write('</sheetData><autoFilter ref="A1:F%d"/>' % (count + 1))
        output.write('</worksheet>')
        sheet_tmp.seek(0)
        z.writestr(name, sheet_tmp.read())
    finally:
        sheet_tmp.close()
    tmp.close()
    return count

def build(opts):
    start = datetime.datetime.strptime(opts.start_date, '%Y-%m-%d').date()
    end = datetime.datetime.strptime(opts.end_date, '%Y-%m-%d').date()
    if end < start:
        raise ValueError('--end-date must be on or after --start-date')
    items = list(periods(start, end, opts.granularity))
    if len(items) > 500:
        raise ValueError('%d worksheets would be needed (max 500). Use a coarser granularity or shorter date range.' % len(items))
    tmp_path = opts.output + '.partial'
    if os.path.exists(tmp_path):
        os.remove(tmp_path)
    try:
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED, allowZip64=True) as z:
            sheet_info = []
            for i, (begin, finish, label) in enumerate(items, 1):
                print('[%d/%d] Exporting %s (%s to %s) ...' % (
                    i, len(items), label, begin, finish))
                sys.stdout.flush()
                rows = write_sheet(z, i, begin, finish, label, opts)
                print('       %d rows' % rows)
                sys.stdout.flush()
                sheet_info.append(label)
            workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                        '<sheets>')
            for i, label in enumerate(sheet_info, 1):
                workbook += '<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (label, i, i)
            workbook += '</sheets></workbook>'
            z.writestr('xl/workbook.xml', workbook)
            rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            for i in range(1, len(sheet_info) + 1):
                rels += '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i, i)
            rels += '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' % (len(sheet_info) + 1)
            z.writestr('xl/_rels/workbook.xml.rels', rels)
            z.writestr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
            types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            for i in range(1, len(sheet_info) + 1):
                types += '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i
            z.writestr('[Content_Types].xml', types + '</Types>')
            z.writestr('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF17365D"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>')
        os.rename(tmp_path, opts.output)
        print('Ready: %s' % os.path.abspath(opts.output))
        print('Note: one row per API product/proxy/app per sheet; error_count is included in message_count.')
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise

def main():
    p = argparse.ArgumentParser(description='Export Apigee analytics into XLSX sheets by hour/day/week/month. Python 2.7; no pip.')
    p.add_argument('--host', default=DEFAULT_HOST)
    p.add_argument('--user', default='apigee')
    p.add_argument('--database', default='apigee')
    p.add_argument('--table', required=True, help='Example: analytics."ORG.ENV.fact"')
    p.add_argument('--start-date', help='Inclusive first day, YYYY-MM-DD')
    p.add_argument('--end-date', help='Inclusive last day, YYYY-MM-DD')
    p.add_argument('--granularity', choices=['hourly', 'daily', 'weekly', 'monthly'],
                   help='One sheet per hour/day/week/month')
    p.add_argument('--output', default='apigee_developer_engagement.xlsx')
    opts = p.parse_args()
    try:
        opts.table = allowed_table(opts.table)
        # Interactive choices when arguments are missing (also works in Python 2.7).
        today = datetime.date.today()
        if not opts.start_date:
            opts.start_date = raw_input('Start date YYYY-MM-DD [%s]: ' % ((today - datetime.timedelta(days=7)).isoformat())).strip() or (today - datetime.timedelta(days=7)).isoformat()
        if not opts.end_date:
            opts.end_date = raw_input('End date YYYY-MM-DD (inclusive) [%s]: ' % ((today - datetime.timedelta(days=1)).isoformat())).strip() or (today - datetime.timedelta(days=1)).isoformat()
        if not opts.granularity:
            selection = raw_input('Bontas: 1=oras, 2=napi, 3=heti, 4=havi [2]: ').strip() or '2'
            opts.granularity = {'1':'hourly','2':'daily','3':'weekly','4':'monthly'}.get(selection)
            if opts.granularity is None:
                raise ValueError('Invalid granularity selection')
        build(opts)
    except (ValueError, RuntimeError, OSError, subprocess.CalledProcessError) as e:
        sys.stderr.write('ERROR: %s\n' % e)
        sys.exit(1)

if __name__ == '__main__':
    main()
