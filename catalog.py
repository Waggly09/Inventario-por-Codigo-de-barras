"""Offline catalog. Codes are strings; identification never uses fuzzy matching."""
from __future__ import annotations
import csv
import json
import math
import os
import re
import sqlite3
import unicodedata
import uuid
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path

FIELDS = ('code','name','category','sku','destination','capacity','alcohol','family','dv','itf14','notes')
LABELS = dict(zip(FIELDS, ('Código base','Producto / descripción','Tipo','SKU','Destino','Capacidad','% Alc. Vol.','Familia','Dígito verificador (DV)','ITF14','Notas')))

def normalize_code(value):
    if isinstance(value, bool) or value is None:
        raise ValueError('El código debe contener solamente dígitos; conserva los ceros iniciales.')
    if isinstance(value, (int, float)):
        if not math.isfinite(value) or int(value)!=value or value<0 or len(str(int(value)))>15:
            raise ValueError('Código numérico inválido o con más de 15 dígitos: guárdalo como texto en Excel.')
        value = str(int(value))
    value = re.sub(r'\s+', '', str(value))
    if not re.fullmatch(r'[0-9]{1,64}', value):
        raise ValueError('Código inválido. Usa de 1 a 64 dígitos, sin guiones, letras ni notación científica.')
    return value

def gs1_dv(base):
    return str((-sum(int(n)*(3 if i%2==0 else 1) for i,n in enumerate(reversed(base))))%10)

def fold(text):
    return ''.join(c for c in unicodedata.normalize('NFD',str(text)).casefold() if unicodedata.category(c)!='Mn')

def validate_record(item):
    if not isinstance(item,dict): raise ValueError('Registro de producto inválido.')
    r = dict(item)
    for key in FIELDS:
        value = r.get(key,'')
        if value is None: value=''
        r[key]=str(value).strip()
    r['code']=normalize_code(item.get('code'))
    if r['category'] not in ('Producto','Corrugado'): raise ValueError('Selecciona Producto o Corrugado.')
    if r['dv'] and not re.fullmatch('[0-9]',r['dv']): raise ValueError('El DV debe ser un solo dígito o quedar vacío.')
    if r['itf14']: r['itf14']=normalize_code(r['itf14'])
    if not isinstance(r.get('extras',{}),dict): raise ValueError('Características adicionales inválidas.')
    if any(len(r[f])>10000 for f in FIELDS): raise ValueError('Un campo supera los 10,000 caracteres.')
    r.setdefault('id',uuid.uuid4().hex)
    if not isinstance(r['id'],str) or not r['id']: raise ValueError('Identificador inválido.')
    return r

def aliases(record):
    yield record['code'],'Código base'
    if record.get('itf14'): yield record['itf14'],'ITF14 en Excel'
    if len(record['code'])==12:
        dv=record.get('dv','')
        yield record['code']+(dv or gs1_dv(record['code'])), 'DV en Excel' if dv else 'DV calculado'

def atomic_json(path, data):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    temp=path.with_name(path.name+'.'+uuid.uuid4().hex+'.tmp')
    try:
        with temp.open('w',encoding='utf-8') as f:
            json.dump(data,f,ensure_ascii=False,indent=2); f.flush();os.fsync(f.fileno())
        os.replace(temp,path)
    finally:
        if temp.exists(): temp.unlink()

def backup_payload(rows):
    return {'format':'corrugados-catalog','version':1,'created':datetime.now(timezone.utc).isoformat(),'records':rows}

def read_backup(path):
    with open(path,encoding='utf-8') as f: data=json.load(f)
    if not isinstance(data,dict) or data.get('format')!='corrugados-catalog' or data.get('version')!=1 or not isinstance(data.get('records'),list):
        raise ValueError('Este archivo no es un respaldo compatible de Corrugados.')
    rows=[validate_record(r) for r in data['records']]
    if len({r['id'] for r in rows})!=len(rows): raise ValueError('El respaldo tiene identificadores repetidos.')
    return rows

def cell_text(cell):
    v=cell.value
    if v is None: return ''
    if isinstance(v,(int,float)):
        if '%' in cell.number_format: return f'{v*100:g}%'
        if int(v)==v:
            text=str(int(v))
            if re.fullmatch(r'0+',cell.number_format): text=text.zfill(len(cell.number_format))
            return text
    return str(v).strip()

def read_excel(path):
    from openpyxl import load_workbook
    w=load_workbook(path,data_only=True,read_only=True)
    rows=[];issues=[];found=0
    mapping={'producto':'name','sku de etiqueta contra':'sku','sku de producto':'sku','destino':'destination','capacidad':'capacity','% alc vol.':'alcohol','familia':'family','dv':'dv','itf14':'itf14'}
    try:
        for s in w:
            header=None
            for row_number,cells in enumerate(s.iter_rows(),1):
                texts=[cell_text(c) for c in cells]
                h=[fold(t).strip() for t in texts]
                if 'numero base + numero movil' in h:
                    header=h;found+=1;continue
                if header is None or not any(texts): continue
                codecol=header.index('numero base + numero movil')
                base=texts[codecol]
                if not base:
                    issues.append(f'{s.title.strip()} · fila {row_number}: falta código; registro omitido.');continue
                r={'category':'Corrugado' if 'dv' in header or 'corrugado' in fold(s.title) else 'Producto','source_sheet':s.title.strip(),'source_row':row_number,'raw_code':base,'extras':{}}
                for ix,(head,text) in enumerate(zip(header,texts)):
                    if ix==codecol: continue
                    if head in mapping: r[mapping[head]]=text
                    elif text: r['extras'][head or f'Columna {ix+1}']=text
                try:
                    # Numeric Excel values beyond its 15-digit precision cannot be trusted.
                    original=cells[codecol].value
                    if isinstance(original,(int,float)): normalize_code(original)
                    r['code']=normalize_code(base)
                    r['id']=uuid.uuid5(uuid.NAMESPACE_URL,f'corrugados/{s.title.strip()}/{row_number}').hex
                    r=validate_record(r)
                except ValueError as e:
                    issues.append(f'{s.title.strip()} · fila {row_number}: {e} Registro omitido.');continue
                if not r['name']: issues.append(f'{s.title.strip()} · fila {row_number}: sin descripción; se conserva el registro.')
                rows.append(r)
    finally: w.close()
    if not found: raise ValueError('No se encontró el encabezado NUMERO BASE + NUMERO MOVIL. Usa el formato del Excel original.')
    if not rows: raise ValueError('El Excel no contiene códigos válidos para importar.\n'+'\n'.join(issues[:8]))
    return rows,issues

class Store:
    def __init__(self,path):
        self.path=Path(path);self.path.parent.mkdir(parents=True,exist_ok=True)
        self.db=sqlite3.connect(self.path,timeout=8)
        self.db.execute('PRAGMA synchronous=FULL')
        self.db.executescript('CREATE TABLE IF NOT EXISTS records(id TEXT PRIMARY KEY,payload TEXT NOT NULL,deleted INTEGER NOT NULL DEFAULT 0,changed TEXT NOT NULL); CREATE TABLE IF NOT EXISTS history(at TEXT,action TEXT,record_id TEXT); PRAGMA user_version=1;')
        self.refresh()

    def refresh(self):
        self.rows=[json.loads(row[0]) for row in self.db.execute('SELECT payload FROM records WHERE deleted=0 ORDER BY rowid')]
        self.index=defaultdict(dict)
        for r in self.rows:
            for code,kind in aliases(r):
                existing=self.index[code].get(r['id'])
                if existing is None or (existing['kind']=='DV calculado' and kind!='DV calculado'):
                    self.index[code][r['id']]={'record':r,'kind':kind}

    def list(self,query='',category='Todos'):
        q=fold(query)
        return [dict(r) for r in self.rows if (category=='Todos' or r['category']==category) and (not q or q in fold(' '.join(str(v) for v in r.values())))]

    def lookup(self,code):
        return sorted(self.index.get(normalize_code(code),{}).values(),key=lambda m:(m['kind']=='DV calculado',m['record']['category'],m['record']['name']))

    def snapshot(self):
        if not self.rows: return
        folder=self.path.parent/'Respaldos'
        stamp=datetime.now().strftime('%Y%m%d-%H%M%S-%f')
        self.export_backup(folder/f'{stamp}.corrugados')
        for old in sorted(folder.glob('*.corrugados'))[:-30]: old.unlink()

    def export_backup(self,path): atomic_json(path,backup_payload(self.rows))

    def replace(self,items):
        records=[validate_record(r) for r in items]
        if len({r['id'] for r in records})!=len(records): raise ValueError('Identificadores duplicados.')
        self.snapshot()
        stamp=datetime.now(timezone.utc).isoformat()
        with self.db:
            self.db.execute('DELETE FROM records')
            self.db.executemany('INSERT INTO records VALUES(?,?,0,?)',[(r['id'],json.dumps(r,ensure_ascii=False),stamp) for r in records])
            self.db.execute('INSERT INTO history VALUES(?,?,?)',(stamp,'Reemplazar catálogo',str(len(records))))
        self.refresh()

    def save(self,item):
        r=validate_record(item); self.snapshot();stamp=datetime.now(timezone.utc).isoformat()
        with self.db:
            self.db.execute('INSERT INTO records VALUES(?,?,0,?) ON CONFLICT(id) DO UPDATE SET payload=excluded.payload,deleted=0,changed=excluded.changed',(r['id'],json.dumps(r,ensure_ascii=False),stamp))
            self.db.execute('INSERT INTO history VALUES(?,?,?)',(stamp,'Guardar',r['id']))
        self.refresh();return r

    def delete(self,id):
        self.snapshot();stamp=datetime.now(timezone.utc).isoformat()
        with self.db:
            self.db.execute('UPDATE records SET deleted=1,changed=? WHERE id=?',(stamp,id))
            self.db.execute('INSERT INTO history VALUES(?,?,?)',(stamp,'Eliminar',id))
        self.refresh()

    def undo_delete(self):
        row=self.db.execute('SELECT id FROM records WHERE deleted=1 ORDER BY changed DESC LIMIT 1').fetchone()
        if not row:return False
        self.snapshot()
        with self.db:
            self.db.execute('UPDATE records SET deleted=0 WHERE id=?',row)
            self.db.execute('INSERT INTO history VALUES(?,?,?)',(datetime.now(timezone.utc).isoformat(),'Recuperar eliminado',row[0]))
        self.refresh();return True

    def report(self):
        counts=Counter(r['code'] for r in self.rows)
        lines=[f'Registros activos: {len(self.rows)}',f'Productos: {sum(r["category"]=="Producto" for r in self.rows)}',f'Corrugados: {sum(r["category"]=="Corrugado" for r in self.rows)}',f'Códigos base compartidos: {sum(v>1 for v in counts.values())}',f'Sin descripción: {sum(not r["name"] for r in self.rows)}','', 'CÓDIGOS BASE COMPARTIDOS']
        for code,count in sorted(counts.items()):
            if count>1:
                lines.append(f'\n{code} — {count} registros')
                for r in self.rows:
                    if r['code']==code: lines.append(f'  {r["category"]}: {r["name"] or "Sin descripción"} ({r.get("source_sheet","Manual")}, fila {r.get("source_row","—")})')
        lines+=['','REVISIÓN DE DÍGITOS VERIFICADORES (no se corrigen automáticamente)']
        for r in self.rows:
            if len(r['code'])==12 and r.get('dv') and r['dv']!=gs1_dv(r['code']): lines.append(f'{r["code"]}: DV registrado {r["dv"]}; cálculo GS1 {gs1_dv(r["code"])}. {r["name"]}')
            if r.get('itf14') and len(r['itf14'])!=14: lines.append(f'{r["code"]}: ITF14 registrado tiene {len(r["itf14"])} dígitos: {r["itf14"]}. Verifica la etiqueta.')
        return '\n'.join(lines)

    def export_csv(self,path):
        with open(path,'w',encoding='utf-8-sig',newline='') as f:
            writer=csv.writer(f,delimiter=';');writer.writerow([LABELS[k] for k in FIELDS])
            for r in self.rows:
                writer.writerow([("'"+r[k]) if k in ('code','dv','itf14') or r[k].startswith(('=','+','-','@','\t','\r')) else r[k] for k in FIELDS])

    def close(self):
        if self.db is not None: self.db.close();self.db=None
