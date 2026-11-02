"""Durable SQLite snapshot storage for scraped vendor assertions and source evidence."""
from __future__ import annotations
import json
import sqlite3
from pathlib import Path
from datetime import datetime, timezone
from typing import Iterable

DEFAULT_DB = Path(__file__).resolve().parent / 'state' / 'compatibility.sqlite3'

def connect(path: str | Path = DEFAULT_DB):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    con=sqlite3.connect(str(path),timeout=15)
    con.row_factory=sqlite3.Row
    con.execute('PRAGMA busy_timeout=15000')
    con.executescript('''
    CREATE TABLE IF NOT EXISTS snapshots (
      id INTEGER PRIMARY KEY AUTOINCREMENT, vendor TEXT NOT NULL, captured_at TEXT NOT NULL,
      row_count INTEGER NOT NULL, stats_json TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS observations (
      snapshot_id INTEGER NOT NULL, row_key TEXT NOT NULL, vendor TEXT NOT NULL,
      product TEXT NOT NULL, vrm TEXT NOT NULL, zos_version TEXT NOT NULL,
      status TEXT NOT NULL, ptfs_json TEXT NOT NULL, links_json TEXT NOT NULL,
      source_url TEXT NOT NULL, evidence_status TEXT NOT NULL, display_name TEXT NOT NULL,
      PRIMARY KEY(snapshot_id,row_key), FOREIGN KEY(snapshot_id) REFERENCES snapshots(id));
    CREATE TABLE IF NOT EXISTS changes (
      snapshot_id INTEGER NOT NULL, row_key TEXT NOT NULL, change_type TEXT NOT NULL,
      before_json TEXT, after_json TEXT, FOREIGN KEY(snapshot_id) REFERENCES snapshots(id));
    CREATE TABLE IF NOT EXISTS jobs (
      id INTEGER PRIMARY KEY AUTOINCREMENT, vendor TEXT, executed_at TEXT, result TEXT,
      message TEXT, records INTEGER DEFAULT 0);
    CREATE INDEX IF NOT EXISTS idx_obs_match ON observations (vendor, product, vrm, zos_version);
    ''')
    return con

def key(r):
    return '|'.join(str(r.get(k,'')).strip().casefold() for k in ('vendor','product','vrm','zos_version'))

def compact(r):
    return {k:r.get(k) for k in ('status','ptfs','links','evidence_status','source_url')}

def save(records: list[dict], stats: dict, path=DEFAULT_DB):
    if not records:raise ValueError('Empty crawler results cannot be published')
    vendor=records[0]['vendor']
    if any(r['vendor']!=vendor for r in records):raise ValueError('One vendor per snapshot')
    with connect(path) as conn:
        prev=conn.execute('SELECT id FROM snapshots WHERE vendor=? ORDER BY id DESC LIMIT 1',(vendor,)).fetchone()
        old={}
        if prev:
            for row in conn.execute('SELECT * FROM observations WHERE snapshot_id=?',(prev['id'],)):
                old[row['row_key']]={'status':row['status'],'ptfs':json.loads(row['ptfs_json']),'links':json.loads(row['links_json']),'evidence_status':row['evidence_status'],'source_url':row['source_url']}
        now=datetime.now(timezone.utc).isoformat()
        cur=conn.execute('INSERT INTO snapshots(vendor,captured_at,row_count,stats_json) VALUES(?,?,?,?)',(vendor,now,len(records),json.dumps(stats)))
        snapshot=cur.lastrowid
        new={}
        for r in records:
            k=key(r);c=compact(r);new[k]=c
            conn.execute('INSERT INTO observations VALUES(?,?,?,?,?,?,?,?,?,?,?,?)',(snapshot,k,r['vendor'],r['product'],r['vrm'],r['zos_version'],r['status'],json.dumps(r['ptfs']),json.dumps(r['links']),r['source_url'],r['evidence_status'],r.get('display_name','')))
            if prev and (k not in old or old[k]!=c):
                conn.execute('INSERT INTO changes VALUES(?,?,?,?,?)',(snapshot,k,'Added' if k not in old else 'Updated',json.dumps(old.get(k)),json.dumps(c)))
        if prev:
            for k,v in old.items():
                if k not in new:conn.execute('INSERT INTO changes VALUES(?,?,?,?,?)',(snapshot,k,'Missing from new source',json.dumps(v),None))
        conn.execute('INSERT INTO jobs(vendor,executed_at,result,message,records) VALUES(?,?,?,?,?)',(vendor,now,'Success','Published validated snapshot',len(records)))
        return snapshot

def record_error(vendor: str, error: str,path=DEFAULT_DB):
    with connect(path) as conn:
        conn.execute('INSERT INTO jobs(vendor,executed_at,result,message,records) VALUES(?,?,?,?,?)',(vendor,datetime.now(timezone.utc).isoformat(),'Failed',str(error)[:800],0))

def latest(vendor='Broadcom',path=DEFAULT_DB):
    with connect(path) as c:
        info=c.execute('SELECT * FROM snapshots WHERE vendor=? ORDER BY id DESC LIMIT 1',(vendor,)).fetchone()
        if not info:return [],None
        records=[]
        for row in c.execute('SELECT * FROM observations WHERE snapshot_id=? ORDER BY product,vrm,zos_version',(info['id'],)):
            r=dict(row);r['ptfs']=json.loads(r.pop('ptfs_json'));r['links']=json.loads(r.pop('links_json'));records.append(r)
        return records,dict(info)

def changes(vendor='Broadcom',limit=200,path=DEFAULT_DB):
    with connect(path) as c:
        return [dict(x) for x in c.execute('''SELECT ch.*,s.captured_at FROM changes ch JOIN snapshots s ON s.id=ch.snapshot_id WHERE s.vendor=? ORDER BY ch.snapshot_id DESC LIMIT ?''',(vendor,limit))]

def jobs(limit=30,path=DEFAULT_DB):
    with connect(path) as c:
        return [dict(x) for x in c.execute('SELECT * FROM jobs ORDER BY id DESC LIMIT ?',(limit,))]
