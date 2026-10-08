#!/usr/bin/env python3
"""Read-only XLSX verification using Python's standard library."""
import argparse
import json
import pathlib
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

APP=pathlib.Path(__file__).resolve().parent
NS={'x':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
ANCHORS={
  'tracker':[['A2','A6','C6','D6','H6','C7','D7','A10','A15','H15'],[f'{c}8' for c in 'ABCDEFGHIJ']+['A2','A9','C9','E9','F9'],[f'{c}6' for c in 'ABCDEFGH']+['A2','G7','H7']],
  'historical':[['A2','A6','C6','D6','H6','C7','D7','A10','A15','H15','A25','B25','C25'],[f'{c}6' for c in 'ABCDEFGHIJKL']+['A2','A7','B7','D7','F7','G7']],
}
HEADER_ROWS={'tracker':[6,8,6],'historical':[6,6]}
WIDTHS={'tracker':[8,10,8],'historical':[8,12]}

def canonical(element):
    if element is None:return None
    return {'tag':element.tag.split('}')[-1],'attrs':dict(sorted(element.attrib.items())),'text':(element.text or '').strip(),'children':[canonical(c) for c in element]}

class Book:
    def __init__(self,filename):
        self.z=zipfile.ZipFile(filename)
        assert self.z.testzip() is None, 'Corrupt XLSX ZIP'
        self.styles=ET.fromstring(self.z.read('xl/styles.xml'))
        self.names=[s.attrib['name'] for s in ET.fromstring(self.z.read('xl/workbook.xml')).findall('x:sheets/x:sheet',NS)]
        self.strings=[]
        if 'xl/sharedStrings.xml' in self.z.namelist():
            self.strings=[''.join(s.itertext()) for s in ET.fromstring(self.z.read('xl/sharedStrings.xml')).findall('x:si',NS)]
        self.sheets=[ET.fromstring(self.z.read(f'xl/worksheets/sheet{i+1}.xml')) for i in range(len(self.names))]
        self.cells=[{c.attrib['r']:c for c in s.findall('x:sheetData/x:row/x:c',NS)} for s in self.sheets]
        self.tables={t.attrib['name']:t for p in self.z.namelist() if p.startswith('xl/tables/') and p.endswith('.xml') for t in [ET.fromstring(self.z.read(p))]}
    def value(self,sheet,address):
        c=self.cells[sheet].get(address)
        if c is None:return ''
        assert c.attrib.get('t')!='e', f'Formula error on {self.names[sheet]}!{address}'
        v=c.find('x:v',NS);value=v.text if v is not None and v.text is not None else ''
        if c.attrib.get('t')=='s':return self.strings[int(value)]
        if c.attrib.get('t')=='inlineStr':return ''.join(c.find('x:is',NS).itertext())
        return value
    def style(self,sheet,address):
        c=self.cells[sheet].get(address)
        index=int(c.attrib.get('s','0')) if c is not None else 0
        xf=self.styles.find('x:cellXfs',NS)[index]
        attrs={k:v for k,v in xf.attrib.items() if k not in ('fontId','fillId','borderId','numFmtId','xfId')}
        number=xf.attrib.get('numFmtId','0')
        numformats={f.attrib['numFmtId']:f.attrib['formatCode'] for f in self.styles.findall('x:numFmts/x:numFmt',NS)}
        return {'attrs':attrs,'font':canonical(self.styles.find('x:fonts',NS)[int(xf.attrib.get('fontId','0'))]),'fill':canonical(self.styles.find('x:fills',NS)[int(xf.attrib.get('fillId','0'))]),'border':canonical(self.styles.find('x:borders',NS)[int(xf.attrib.get('borderId','0'))]),'numberFormat':numformats.get(number,number),'alignment':canonical(xf.find('x:alignment',NS))}
    def layout(self,kind):
        result={'sheetNames':self.names,'sheets':[]}
        for i,sheet in enumerate(self.sheets):
            columns={}
            for c in sheet.findall('x:cols/x:col',NS):
                for n in range(int(c.attrib['min']),int(c.attrib['max'])+1):columns[n]=c.attrib.get('width')
            view=sheet.find('x:sheetViews/x:sheetView',NS);pane=view.find('x:pane',NS)
            rows={r.attrib['r']:r.attrib.get('ht') for r in sheet.findall('x:sheetData/x:row',NS)}
            result['sheets'].append({'columns':[columns.get(n) for n in range(1,WIDTHS[kind][i]+1)],'gridLines':view.attrib.get('showGridLines','1'),'freeze':{k:pane.attrib[k] for k in ('xSplit','ySplit','state') if k in pane.attrib} if pane is not None else None,'headerHeight':rows.get(str(HEADER_ROWS[kind][i])),'titleHeight':rows.get('2'),'merges':[m.attrib['ref'] for m in sheet.findall('x:mergeCells/x:mergeCell',NS)],'styles':{a:self.style(i,a) for a in ANCHORS[kind][i]}})
        return result

def assert_close(actual,expected,label):
    if expected is None:assert actual=='—', f'{label}: expected unavailable rate, got {actual}'
    else:assert abs(float(actual)-expected)<1e-10, f'{label}: {actual} versus {expected}'

def verify(args):
    books={'tracker':Book(args.tracker),'historical':Book(args.historical)}
    reference=json.loads((APP/'layout_reference.json').read_text())
    format=json.loads((APP/'format.json').read_text())
    for kind,book in books.items():
        assert book.layout(kind)==reference[kind], f'{kind}: columns, fonts, colors, frozen panes, or fixed positions differ from the approved layout.'
        assert book.names==format['workbooks'][kind]
        for i,cells in enumerate(book.cells):
            for address in cells:book.value(i,address)
        assert not any('/charts/' in p for p in book.z.namelist()), 'This format contains no conversion-rate charts.'
    expected_tables={
      'tracker':{'FinancePlacementTracker':'tracker','FinancePlacementReview':'review','FinanceTrackerStatusSummary':'trackerStatus','FinanceTrackerInformationSummary':'trackerCoverage'},
      'historical':{'FinanceHistoricalStudents':'historicalDetail','FinanceHistoricalStatusSummary':'historicalStatus','FinanceHistoricalCoverage':'historicalCoverage','FinanceInternshipEmployerCounts':'internshipEmployerCounts'}
    }
    for kind,tables in expected_tables.items():
        assert set(books[kind].tables)==set(tables), f'{kind}: unexpected or missing tables.'
        for name,header_key in tables.items():
            table=books[kind].tables[name]
            headings=[c.attrib['name'] for c in table.find('x:tableColumns',NS)]
            assert headings==format['headers'][header_key], f'{name}: column headings changed.'
            assert table.find('x:autoFilter',NS) is not None
    if not args.model:return
    model=json.loads(args.model)
    for kind in books:
        book=books[kind];stats=model[kind]
        for r,expected in zip((7,8,9),stats):
            values=[int(book.value(0,f'{c}{r}')) for c in ('B','C','E','F','G')]
            assert values==[expected[k] for k in ('students','placed','activelySeeking','notSeeking','noStatus')]
            assert sum(values[1:])==values[0]
            assert_close(book.value(0,f'D{r}'),expected['placed']/expected['students'] if expected['students'] else None,f'{kind} rate')
        assert int(book.value(0,'B10'))==sum(s['students'] for s in stats)
    for kind,table_name,id_col,cohort_col in [('tracker','FinancePlacementTracker','A','G'),('historical','FinanceHistoricalStudents','B','A')]:
        book=books[kind];ref=book.tables[table_name].attrib['ref'];first,last=map(int,re.findall(r'\d+',ref))
        ids=[book.value(1,f'{id_col}{r}') for r in range(first+1,last+1) if book.value(1,f'{id_col}{r}')]
        assert len(ids)==len(set(ids))==sum(s['students'] for s in model[kind])
        allowed={str(s['term'] if kind=='tracker' else s['year']) for s in model[kind]}
        assert {book.value(1,f'{cohort_col}{r}') for r in range(first+1,last+1) if book.value(1,f'{id_col}{r}')}<=allowed
    print('Verified fixed layout, headers, unique scoped students, cached formulas, totals, rates, filters, and freeze panes.')

if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--tracker',required=True);parser.add_argument('--historical',required=True);parser.add_argument('--model')
    args=parser.parse_args()
    try:verify(args)
    except Exception as error:print(f'Verification failed: {error}',file=sys.stderr);sys.exit(1)
