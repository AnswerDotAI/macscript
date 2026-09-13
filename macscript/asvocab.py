"Ergonomic exploration of an appscript app's scripting vocabulary: live terminology tables, plus the sdef XML for full docs."
import re
from functools import cache
from aem import AEEnum
from xml.etree import ElementTree as etree

__all__ = ['vocab', 'props', 'sd', 'sdfind']

_kinds = {b'p': 'property', b'e': 'element', b'c': 'command'}

def vocab(app, pat=''):
    "Search `app`'s terminology for names matching regex `pat` (case-insensitive), as sorted (kind, name) rows"
    r = re.compile(pat, re.I)
    ad = app.AS_appdata
    res = [(_kinds[v[0]], n) for n,v in ad.referencebyname().items() if r.search(n)]
    res += [('enum' if isinstance(v, AEEnum) else 'class', n) for n,v in ad.typebyname().items() if r.search(n)]
    return sorted(res)

def props(ref, pat=None, maxlen=80, timeout=15):
    "`ref.properties.get()` as a readable {name: truncated-repr} dict, optionally filtered by regex `pat`"
    r = re.compile(pat, re.I) if pat else None
    res = {}
    for key, v in ref.properties.get(timeout=timeout).items():
        n = str(key).removeprefix('k.')
        if r and not r.search(n): continue
        s = repr(v)
        res[n] = s if len(s) <= maxlen else s[:maxlen] + '...'
    return res

class _Doc(str):
    def __repr__(self): return str(self)

@cache
def _sdef(path): return etree.parse(str(path))

def _typ(e):
    if e.get('type'): return e.get('type')
    return ' | '.join(('list of ' if t.get('list') else '')+t.get('type','') for t in e.findall('type'))

def _desc(e): return f' -- {e.get("description")}' if e.get('description') else ''

def _fmt_cmd(c):
    lines = [f'command {c.get("name")}{_desc(c)}']
    dp = c.find('direct-parameter')
    if dp is not None: lines.append(f'  direct: {_typ(dp)}{_desc(dp)}')
    for p in c.findall('parameter'):
        opt = ' (optional)' if p.get('optional')=='yes' else ''
        lines.append(f'  {p.get("name")}: {_typ(p)}{opt}{_desc(p)}')
    r = c.find('result')
    if r is not None: lines.append(f'  result: {_typ(r)}{_desc(r)}')
    return lines

def _fmt_cls(c):
    inh = f' < {c.get("inherits")}' if c.get('inherits') else ''
    lines = [f'class {c.get("name")}{inh}{_desc(c)}']
    for e in c.findall('element'): lines.append(f'  element: {e.get("type")}')
    for p in c.findall('property'):
        acc = ' (r/o)' if p.get('access')=='r' else ''
        lines.append(f'  {p.get("name")}: {p.get("type")}{acc}{_desc(p)}')
    return lines

def _fmt_enum(c):
    lines = [f'enumeration {c.get("name")}']
    for e in c.findall('enumerator'): lines.append(f'  {e.get("name")}{_desc(e)}')
    return lines

_fmts = {'command': _fmt_cmd, 'class': _fmt_cls, 'enumeration': _fmt_enum}
_children = ('parameter', 'property', 'enumerator')

def sd(name, path):
    "Full sdef doc for `name` (a command, class, or enumeration; underscore and space spellings both work), with types and descriptions"
    n = name.replace('_', ' ')
    nodes = [e for tag in _fmts for e in _sdef(path).iter(tag) if e.get('name') in (n, name)]
    if not nodes: raise KeyError(f'{name!r} not found in sdef; try sdfind')
    return _Doc('\n\n'.join('\n'.join(_fmts[e.tag](e)) for e in nodes))

def sdfind(pat, path, maxlen=80):
    "Search sdef names and descriptions for regex `pat`, as (kind, name, description) rows; child nodes show as parent.name"
    r = re.compile(pat, re.I)
    res = []
    root = _sdef(path)
    parents = {c: p for p in root.iter() for c in p}
    for e in root.iter():
        if e.tag not in _fmts and e.tag not in _children: continue
        n, d = e.get('name') or '', e.get('description') or ''
        if not (r.search(n) or r.search(d)): continue
        if e.tag in _children: n = f'{parents[e].get("name")}.{n}'
        res.append((e.tag, n, d if len(d)<=maxlen else d[:maxlen]+'...'))
    return res
