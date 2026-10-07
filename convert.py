#!/usr/bin/env python3
"""Convert employee workbook to roster JSON, optionally merging existing bookings."""
import argparse, datetime as dt, json, re, sys
from pathlib import Path

def clean(v):
    return str(v).strip() if v is not None else ''
def norm(v):
    return re.sub(r'[^a-z0-9]', '', clean(v).lower())
def convert(source, existing=None):
    try:
        import openpyxl
    except ImportError:
        raise SystemExit('Install Excel support: python -m pip install openpyxl')
    book=openpyxl.load_workbook(source, data_only=True)
    data=existing or {'schemaVersion':1,'settings':{'title':'Service Team Roster','workDays':[1,2,3,4,5]},'employees':[],'clients':[],'projects':[],'assignments':[],'leaves':[]}
    lookup={norm(e['name']):e for e in data['employees']}
    aliases={'name':{'empname','employeename','name'},'role':{'jobtitle','designation','role'},'department':{'department'},'joiningDate':{'joiningdate','dateofjoining'},'type':{'type'},'company':{'visaissuancecompany','company'},'skills':{'skills','skillset'}}
    found=set()
    for sheet in book:
        header=None
        for row in sheet.iter_rows(values_only=True):
            if header is None:
                cols={k:next((i for i,v in enumerate(row) if norm(v) in names),None) for k,names in aliases.items()}
                if cols['name'] is not None:header=cols
                continue
            name=clean(row[header['name']])
            if not name or norm(name) in aliases['name']:continue
            key=norm(name)
            if key in found:raise ValueError(f'Duplicate employee name: {name}. Resolve duplicate names before converting.')
            found.add(key)
            emp=lookup.get(key)
            if emp is None:
                used={e['id'] for e in data['employees']}
                i=1
                while f'EMP-{i:03d}' in used:i+=1
                emp={'id':f'EMP-{i:03d}','active':True,'capacity':100,'notes':''}
                data['employees'].append(emp)
            for field,index in header.items():
                if index is None:continue
                value=row[index]
                if field=='joiningDate':
                    if isinstance(value,(dt.datetime,dt.date)):value=value.date().isoformat() if isinstance(value,dt.datetime) else value.isoformat()
                    elif value:
                        try:value=dt.date.fromisoformat(clean(value)).isoformat()
                        except ValueError:raise ValueError(f'Invalid joining date for {name}: {value}. Use Excel date cells or YYYY-MM-DD.')
                emp[field]=clean(value)
            emp.setdefault('skills',emp.get('role',''))
    if not found:raise ValueError('No employee header found. Expected EMP. NAME or Employee Name.')
    data['source']={'filename':Path(source).name,'convertedAt':dt.datetime.now(dt.timezone.utc).isoformat(),'importedEmployees':len(found)}
    return data

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('excel',nargs='?',help='Source .xlsx workbook')
    p.add_argument('-o','--output',default='data.json')
    p.add_argument('--merge',help='Existing exported JSON; preserves assignments, IDs and employees absent from Excel')
    p.add_argument('--serve',action='store_true',help='Serve this folder locally; does not modify files')
    p.add_argument('--port',type=int,default=8000)
    a=p.parse_args()
    if a.serve:
        import http.server
        import os
        os.chdir(Path(__file__).resolve().parent)
        print(f'Open http://localhost:{a.port} (Ctrl+C to stop)')
        http.server.ThreadingHTTPServer(('127.0.0.1',a.port),http.server.SimpleHTTPRequestHandler).serve_forever()
        return
    if not a.excel:p.error('Provide an Excel file, or use --serve')
    existing=json.loads(Path(a.merge).read_text(encoding='utf-8')) if a.merge else None
    if existing is not None:
        if existing.get('schemaVersion')!=1 or any(not isinstance(existing.get(k),list) for k in ['employees','clients','projects','assignments','leaves']):p.error('Merge file is not a roster schema v1 JSON')
    out=convert(a.excel,existing)
    Path(a.output).write_text(json.dumps(out,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    print(f"Saved {len(out['employees'])} employees to {a.output}. Existing bookings preserved with --merge.")
if __name__=='__main__':
    try:main()
    except (ValueError,OSError) as exc:sys.exit(str(exc))
