"""OOXML source extraction/audit and notes-to-DOCX; standard library core."""
import argparse
import hashlib
import importlib.util
import importlib
import json
import posixpath
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as E
from xml.sax.saxutils import escape

NS = {'p':'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'm':'http://schemas.openxmlformats.org/officeDocument/2006/math',
      'rel':'http://schemas.openxmlformats.org/package/2006/relationships'}


def pdf_engine():
    for name in ['pymupdf','fitz']:
        if importlib.util.find_spec(name):
            return importlib.import_module(name)
    raise ValueError('PyMuPDF missing; ask user before installation')


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def text(node, xpath, ns=NS):
    return ''.join(n.text or '' for n in node.findall(xpath, ns))


def rels(z, part):
    name = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
    if name not in z.namelist():
        return {}
    return {x.get('Id'): x for x in E.fromstring(z.read(name))}


def target(part, relationship):
    value = relationship.get('Target', '')
    if value.startswith('/'):
        return value.lstrip('/')
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), value))


def slide_paths(z):
    root = E.fromstring(z.read('ppt/presentation.xml'))
    rs = rels(z, 'ppt/presentation.xml')
    return [target('ppt/presentation.xml', rs[x.get('{'+NS['r']+'}id')])
            for x in root.findall('p:sldIdLst/p:sldId', NS)]


def ppt_slides(path):
    result = []
    with zipfile.ZipFile(path) as z:
        for part in slide_paths(z):
            root = E.fromstring(z.read(part))
            texts = []
            title = ''
            for shape in root.findall('.//p:sp', NS):
                value = '\n'.join(text(p, './/a:t') for p in shape.findall('p:txBody/a:p', NS)).strip()
                if value:
                    texts.append(value)
                name = shape.find('p:nvSpPr/p:cNvPr', NS)
                ph = shape.find('p:nvSpPr/p:nvPr/p:ph', NS)
                if name is not None and name.get('name') == 'PageTitle':
                    title = value
                elif not title and ph is not None and ph.get('type') in {'title','ctrTitle'}:
                    title = value
            if not title:
                title = texts[0] if texts else ''
            notes = ''
            for r in rels(z, part).values():
                if r.get('Type', '').endswith('/notesSlide'):
                    nr = E.fromstring(z.read(target(part, r)))
                    bodies = []
                    for shape in nr.findall('.//p:sp', NS):
                        ph = shape.find('p:nvSpPr/p:nvPr/p:ph', NS)
                        if ph is not None and ph.get('type') == 'body':
                            bodies.extend(text(p, './/a:t') for p in shape.findall('p:txBody/a:p', NS))
                    notes = '\n'.join(bodies).strip()
            result.append({'part':part, 'title':title, 'text':texts, 'notes':notes})
    return result


def docx_paragraphs(path):
    with zipfile.ZipFile(path) as z:
        r = E.fromstring(z.read('word/document.xml'))
        return [word_text(p) for p in r.findall('.//w:p', NS)]


def word_text(node):
    parts=[]
    for child in node.iter():
        if child.tag in {'{'+NS['w']+'}t','{'+NS['m']+'}t'}:
            parts.append(child.text or '')
        elif child.tag=='{'+NS['w']+'}br':
            parts.append('\n')
        elif child.tag=='{'+NS['w']+'}tab':
            parts.append('\t')
    return ''.join(parts)


def audit_pptx(path):
    errors = []
    warnings = []
    counts = {'text_shapes':0,'pictures':0,'tables':0,'charts':0,'connectors':0,'animations':0,'transitions':0}
    external = []
    fonts = set()
    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        if bad:
            errors.append('ZIP CRC error: '+bad)
        names = set(z.namelist())
        for n in names:
            if not n.endswith(('.xml','.rels')):
                continue
            try:
                root = E.fromstring(z.read(n))
            except E.ParseError as e:
                errors.append(f'{n}: {e}'); continue
            if n.endswith('.rels'):
                parent = posixpath.dirname(n)
                part = posixpath.join(posixpath.dirname(parent), posixpath.basename(n)[:-5])
                if n == '_rels/.rels':
                    part = ''
                for r in root:
                    if r.get('TargetMode') == 'External':
                        external.append({'part':part,'target':r.get('Target')})
                    elif target(part, r).split('#')[0] not in names:
                        errors.append(f'Broken relationship {n}: {r.get("Target")}')
        pres = E.fromstring(z.read('ppt/presentation.xml'))
        size = pres.find('p:sldSz', NS)
        width, height = int(size.get('cx')), int(size.get('cy'))
        for n in slide_paths(z):
            root = E.fromstring(z.read(n))
            counts['text_shapes'] += len(root.findall('.//p:sp/p:txBody', NS))
            counts['pictures'] += len(root.findall('.//p:pic', NS))
            counts['tables'] += len(root.findall('.//a:tbl', NS))
            counts['connectors'] += len(root.findall('.//p:cxnSp', NS))
            counts['charts'] += sum(x.tag.endswith('}chart') for x in root.iter())
            counts['animations'] += len(root.findall('p:timing', NS))
            counts['transitions'] += len(root.findall('p:transition', NS))
            for r in root.findall('.//a:rPr', NS):
                for tag in ['latin','ea']:
                    f = r.find('a:'+tag, NS)
                    if f is not None and f.get('typeface'):
                        fonts.add(f.get('typeface'))
            for xf in root.findall('.//p:sp/p:spPr/a:xfrm', NS) + root.findall('.//p:pic/p:spPr/a:xfrm', NS):
                off, ext = xf.find('a:off',NS), xf.find('a:ext',NS)
                if off is None or ext is None:
                    continue
                x,y,cx,cy = (int(off.get('x')),int(off.get('y')),int(ext.get('cx')),int(ext.get('cy')))
                if min(cx,cy) <= 0 or x < -12700 or y < -12700 or x+cx > width+12700 or y+cy > height+12700:
                    warnings.append(f'{n}: object outside canvas or nonpositive dimensions')
    slides = ppt_slides(path)
    for i,s in enumerate(slides,1):
        if not s['notes']:
            errors.append(f'slide {i}: missing speaker notes')
    if counts['animations'] or counts['transitions']:
        errors.append('Deck contains timing/transition elements; required static deck')
    return {'source':str(Path(path).resolve()),'source_sha256':sha(path),
            'slide_count':len(slides),'canvas_pt':[width/12700,height/12700],
            'counts':counts,'fonts':sorted(fonts),'external_relationships':external,
            'errors':errors,'warnings':warnings,
            'not_proven':['factual accuracy','text fit','image readability','visual quality','Office compatibility']}


def notes_docx(source, output):
    out = Path(output)
    if out.exists():
        raise ValueError('Output exists; use a new filename')
    slides = ppt_slides(source)
    if not slides or any(not s['notes'] for s in slides):
        raise ValueError('Final PPT has missing notes')
    def paragraph(value, heading=False):
        # OOXML does not permit raw XML control characters.
        if any(ord(c)<32 and c not in '\n\t\r' for c in value):
            raise ValueError('Illegal control character in notes')
        pr = '<w:keepNext/><w:spacing w:before="240" w:after="120"/>' if heading else '<w:spacing w:after="120" w:line="360" w:lineRule="auto"/>'
        bold = '<w:b/>' if heading else ''
        sz = '28' if heading else '23'
        runs = '<w:br/>'.join('<w:t xml:space="preserve">'+escape(x)+'</w:t>' for x in value.split('\n'))
        return f'<w:p><w:pPr>{pr}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="微软雅黑" w:hAnsi="Microsoft YaHei"/>{bold}<w:color w:val="000000"/><w:sz w:val="{sz}"/></w:rPr>{runs}</w:r></w:p>'
    body = ''.join(paragraph(f'第{i}页  {s["title"]}',True) + paragraph(s['notes']) for i,s in enumerate(slides,1))
    document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="'+NS['w']+'"><w:body>'+body+'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"/></w:sectPr></w:body></w:document>'
    content = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'
    rel = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="'+NS['rel']+'"><Relationship Id="rId1" Type="'+NS['r']+'/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>'
    core = '<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:description>Source PPT SHA-256: '+sha(source)+'</dc:description></cp:coreProperties>'
    out.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        for n,v in {'[Content_Types].xml':content,'_rels/.rels':rel,'word/document.xml':document,'docProps/core.xml':core}.items():
            z.writestr(n,v)
    return {'output':str(out.resolve()),'source_sha256':sha(source),'pages':len(slides),'contains_preview_images':False}


def extract(source, output):
    source, output = Path(source), Path(output)
    if output.exists():
        raise ValueError('Extraction directory exists; choose a new directory')
    output.mkdir(parents=True)
    result = {'source':str(source.resolve()),'sha256':sha(source),'reading_completed':False,
              'note':'Extraction is not full reading; inspect pages, figures, captions and equations.'}
    if source.suffix.lower() in {'.docx','.pptx'}:
        with zipfile.ZipFile(source) as z:
            if source.suffix.lower()=='.docx':
                result['paragraphs']=docx_paragraphs(source)
                root=E.fromstring(z.read('word/document.xml'))
                result['tables']=[[[word_text(c) for c in row.findall('w:tc',NS)] for row in table.findall('w:tr',NS)] for table in root.findall('.//w:tbl',NS)]
                result['math_objects']=len(root.findall('.//m:oMath',NS))
                result['tracked_insertions']=len(root.findall('.//w:ins',NS))
                result['tracked_deletions']=len(root.findall('.//w:del',NS))
                result['auxiliary_notes']={}
                for extra in ['word/footnotes.xml','word/endnotes.xml','word/comments.xml']:
                    if extra in z.namelist():
                        xr=E.fromstring(z.read(extra))
                        result['auxiliary_notes'][extra]=[word_text(p) for p in xr.findall('.//w:p',NS)]
                if result['tracked_insertions'] or result['tracked_deletions']:
                    result['revision_warning']='Tracked changes present: inspect the intended document version with the user; extraction is not acceptance of revisions.'
                if result['math_objects']:
                    (output/'document-with-equations.xml').write_bytes(z.read('word/document.xml'))
            else:
                result['slides']=ppt_slides(source)
            media=[]
            for n in z.namelist():
                if n.startswith(('word/media/','ppt/media/')) and not n.endswith('/'):
                    dest=output/'media'/Path(n).name;dest.parent.mkdir(exist_ok=True)
                    dest.write_bytes(z.read(n));media.append({'package_path':n,'file':str(dest.resolve()),'sha256':sha(dest)})
            result['media']=media
    elif source.suffix.lower()=='.pdf':
        if importlib.util.find_spec('pymupdf') or importlib.util.find_spec('fitz'):
            fitz=pdf_engine()
            with fitz.open(source) as doc:
                result['pages']=[{'page':i+1,'text':page.get_text()} for i,page in enumerate(doc)]
                result['render_command']='Use pdf-render to inspect all actual pages; extraction alone is insufficient.'
        elif importlib.util.find_spec('pypdf'):
            from pypdf import PdfReader
            result['pages']=[{'page':i+1,'text':p.extract_text()} for i,p in enumerate(PdfReader(source).pages)]
        else:
            raise ValueError('PDF reader missing: propose PyMuPDF/pypdf installation and get consent; nothing installed')
    else:
        raise ValueError('Supported: DOCX/PDF/PPTX')
    (output/'source.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    return result


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    sub=parser.add_subparsers(dest='cmd',required=True)
    p=sub.add_parser('audit');p.add_argument('source');p.add_argument('--json')
    for cmd in ['extract','notes-docx','pdf-render']:
        p=sub.add_parser(cmd);p.add_argument('source');p.add_argument('--output',required=True)
    a=parser.parse_args()
    if a.cmd=='audit':
        result=audit_pptx(a.source)
        if a.json:
            Path(a.json).write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps(result,ensure_ascii=False,indent=2))
        return 1 if result['errors'] else 0
    if a.cmd=='notes-docx':
        result=notes_docx(a.source,a.output)
    elif a.cmd=='extract':
        result=extract(a.source,a.output)
        result={'output':str(Path(a.output).resolve()),'reading_completed':False,'source_sha256':result['sha256']}
    else:
        fitz=pdf_engine()
        output=Path(a.output)
        if output.exists():
            raise ValueError('Output exists')
        output.mkdir(parents=True)
        with fitz.open(a.source) as doc:
            for i,page in enumerate(doc,1):
                page.get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(output/f'page-{i:03}.png')
        result={'output':str(output.resolve()),'source_sha256':sha(a.source)}
    print(json.dumps(result,ensure_ascii=False,indent=2));return 0


if __name__=='__main__':
    try:
        sys.exit(main())
    except (OSError,ValueError,KeyError,zipfile.BadZipFile,E.ParseError) as e:
        print(f'ERROR: {e}',file=sys.stderr);sys.exit(2)
