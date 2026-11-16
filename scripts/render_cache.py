"""Cache PowerPoint previews by slide dependency graph, renderer and font environment.

Standard library only. A cache hit reuses rendering, not proof of human review.
Notes are excluded; slide geometry, images, layouts, masters and themes are included.
"""
import argparse
import hashlib
import json
import os
import platform
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as E

from office import NS, rels, target, slide_paths, sha

SCHEMA = 'powerpoint-preview-cache-v2'


def visual_xml(z, name, presentation=False):
    """Normalize nonvisual UUIDs and relationship IDs, keeping their actual targets."""
    data = z.read(name)
    if not name.endswith('.xml'):
        return data
    try:
        root = E.fromstring(data)
    except E.ParseError:
        return data
    relationships = rels(z, name)
    for parent in root.iter():
        for child in list(parent):
            if child.tag in {
                '{http://schemas.microsoft.com/office/drawing/2014/main}creationId',
                '{http://schemas.microsoft.com/office/powerpoint/2010/main}creationId'}:
                parent.remove(child)
        if parent.tag.endswith('}cNvPr'):
            for key in ('name', 'descr', 'title'):
                parent.attrib.pop(key, None)
        for key, value in list(parent.attrib.items()):
            if key.startswith('{'+NS['r']+'}') and value in relationships:
                r = relationships[value]
                dest = r.get('Target') if r.get('TargetMode') == 'External' else target(name, r)
                parent.set(key, digest([r.get('Type'), dest, r.get('TargetMode')]))
    if presentation:
        for child in list(root):
            if child.tag in {'{'+NS['p']+'}sldIdLst', '{'+NS['p']+'}notesMasterIdLst'}:
                root.remove(child)
    return E.tostring(root)


def digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False,
                                    separators=(',', ':')).encode('utf-8')).hexdigest()


def font_environment():
    entries = []
    for base in [Path(os.environ.get('WINDIR', 'C:/Windows'))/'Fonts',
                 Path(os.environ.get('LOCALAPPDATA', ''))/'Microsoft/Windows/Fonts']:
        if base.is_dir():
            for p in sorted(base.rglob('*')):
                if p.is_file():
                    s = p.stat()
                    entries.append((str(p), s.st_size, s.st_mtime_ns))
    return digest({'platform': platform.platform(), 'fonts': entries})


def fingerprints(source):
    """Conservative fingerprints; no reuse for decks with external media/objects."""
    with zipfile.ZipFile(source) as z:
        names = set(z.namelist())
        common = hashlib.sha256(visual_xml(z, 'ppt/presentation.xml', True)).hexdigest()
        out = []
        for index, part in enumerate(slide_paths(z), 1):
            seen, entries, cacheable = set(), [], True

            def walk(name):
                nonlocal cacheable
                if name in seen:
                    return
                seen.add(name)
                if name not in names:
                    raise ValueError('Missing visual dependency: '+name)
                data = visual_xml(z, name)
                entries.append((name, hashlib.sha256(data).hexdigest()))
                for r in sorted(rels(z, name).values(), key=lambda r: (r.get('Type', ''), r.get('Target', ''))):
                    kind = r.get('Type', '').rsplit('/', 1)[-1]
                    if kind in {'notesSlide', 'notesMaster', 'commentAuthors', 'comments'}:
                        continue
                    entries.append((name, kind, r.get('Target'), r.get('TargetMode')))
                    if r.get('TargetMode') == 'External':
                        if kind != 'hyperlink':
                            cacheable = False
                    elif kind != 'slide':
                        walk(target(name, r))

            walk(part)
            out.append({'page': index, 'visual_sha256': digest([common, index, entries]),
                        'cacheable': cacheable})
        return out


def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def write(path, value):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')


def plan(source, output, cache, renderer, pages=None):
    source, output, cache = Path(source).resolve(), Path(output).resolve(), Path(cache).resolve()
    output.mkdir(parents=True, exist_ok=True)
    fs = fingerprints(source)
    requested = list(range(1, len(fs)+1)) if not pages else list(dict.fromkeys(pages))
    if not requested or any(i < 1 or i > len(fs) for i in requested):
        raise ValueError('Invalid requested slide numbers')
    if any((output/f'slide-{i:03}.png').exists() for i in requested):
        raise ValueError('Preview files already exist; choose a new output directory')
    jobs = []
    for f in fs:
        if f['page'] not in requested:
            continue
        key = digest([SCHEMA, renderer, f['visual_sha256']])
        meta, png = cache/(key+'.json'), cache/(key+'.png')
        hit, reviewed = False, False
        if f['cacheable'] and meta.is_file() and png.is_file():
            try:
                m = read(meta)
                hit = m.get('key') == key and m.get('sha256') == sha(png)
                reviewed = hit and m.get('visual_reviewed') is True
            except (ValueError, OSError):
                pass
        filename = f"slide-{f['page']:03}.png"
        if hit:
            shutil.copy2(png, output/filename)
        jobs.append({**f, 'file': filename, 'cache_key': key, 'cache_hit': hit,
                     'visual_reviewed': reviewed, 'visual_review_reused': reviewed})
    result = {'schema': SCHEMA, 'source': str(source), 'source_sha256': sha(source),
              'renderer': renderer, 'cache_dir': str(cache), 'slide_count': len(fs),
              'complete_deck': len(requested) == len(fs), 'jobs': jobs}
    write(output/'cache-plan.json', result)
    return result


def complete(report):
    path = Path(report).resolve()
    r = read(path)
    if r.get('error') or not r.get('opened') or sha(r['source']) != r['source_sha256']:
        raise ValueError('Render failed or source changed; cache not committed')
    cache = Path(r['cache_dir'])
    cache.mkdir(parents=True, exist_ok=True)
    for s in r['slides']:
        p = path.parent/s['file']
        if p.read_bytes()[:8] != b'\x89PNG\r\n\x1a\n' or sha(p) != s['sha256']:
            raise ValueError('Missing/corrupt rendered PNG')
        if not s['cacheable']:
            continue
        key = s['cache_key']
        shutil.copy2(p, cache/(key+'.png'))
        write(cache/(key+'.json'), {'schema': SCHEMA, 'key': key, 'sha256': s['sha256'],
                                   'visual_reviewed': s.get('visual_reviewed', False)})
    r['exported_pages'] = [s['page'] for s in r['slides'] if not s['cache_hit']]
    r['reused_pages'] = [s['page'] for s in r['slides'] if s['cache_hit']]
    write(path, r)
    return {'exported_pages': r['exported_pages'], 'reused_pages': r['reused_pages']}


def review(report, pages, note):
    if not note.strip():
        raise ValueError('Record the actual visual inspection, not an empty approval')
    path = Path(report).resolve()
    r = read(path)
    if sha(r['source']) != r['source_sha256']:
        raise ValueError('Report belongs to an older PPTX')
    available = {s['page'] for s in r['slides']}
    if not pages or not set(pages) <= available:
        raise ValueError('Review page is not in this render report')
    for s in r['slides']:
        if s['page'] in pages:
            if sha(path.parent/s['file']) != s['sha256']:
                raise ValueError('Preview changed after render')
            s['visual_reviewed'] = True
            s['review_note'] = note
            meta = Path(r['cache_dir'])/(s['cache_key']+'.json')
            if s['cacheable'] and meta.is_file():
                m = read(meta)
                if m.get('sha256') != s['sha256']:
                    raise ValueError('Cache entry mismatch')
                m.update(visual_reviewed=True, review_note=note)
                write(meta, m)
    write(path, r)
    return {'reviewed_pages': pages, 'human_truth_not_proven': True}


def main():
    p = argparse.ArgumentParser(description=__doc__)
    sub = p.add_subparsers(dest='cmd', required=True)
    a = sub.add_parser('plan')
    for n in ['source', 'output', 'cache']:
        a.add_argument('--'+n, required=True)
    a.add_argument('--version', required=True)
    a.add_argument('--width', type=int, required=True)
    a.add_argument('--height', type=int, required=True)
    a.add_argument('--pages', nargs='*', type=int)
    a = sub.add_parser('complete'); a.add_argument('--report', required=True)
    a = sub.add_parser('review'); a.add_argument('--report', required=True)
    a.add_argument('--pages', nargs='+', type=int, required=True)
    a.add_argument('--note', required=True)
    args = p.parse_args()
    if args.cmd == 'plan':
        result = plan(args.source, args.output, args.cache,
                      {'application': 'PowerPoint', 'version': args.version, 'width': args.width,
                       'height': args.height, 'font_environment': font_environment()}, args.pages)
        print(json.dumps({'export': [j['page'] for j in result['jobs'] if not j['cache_hit']],
                          'reuse': [j['page'] for j in result['jobs'] if j['cache_hit']]}))
    elif args.cmd == 'complete':
        print(json.dumps(complete(args.report)))
    else:
        print(json.dumps(review(args.report, args.pages, args.note), ensure_ascii=False))


if __name__ == '__main__':
    try:
        main()
    except (ValueError, OSError, KeyError, zipfile.BadZipFile) as e:
        raise SystemExit('ERROR: '+str(e))
