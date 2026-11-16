"""Check named native reference-layout anchors; does not replace visual inspection."""
import argparse,json,zipfile
from pathlib import Path
from xml.etree import ElementTree as E
from office import NS,slide_paths


def audit(source, roles):
    findings=[]
    with zipfile.ZipFile(source) as z:
        pages=slide_paths(z)
        if len(pages)!=len(roles):raise ValueError('Role count does not match slide count')
        for i,(part,role) in enumerate(zip(pages,roles),1):
            root=E.fromstring(z.read(part));shapes={s.find('p:nvSpPr/p:cNvPr',NS).get('name'):s for s in root.findall('.//p:sp',NS)}
            def require(name):
                if name not in shapes:findings.append({'page':i,'role':role,'missing':name})
            if role=='contents':
                require('ContentsBluePanel');require('ContentsEnglish')
                panel=shapes.get('ContentsBluePanel')
                if panel is not None:
                    pos=panel.find('p:spPr/a:xfrm',NS)
                    width=int(pos.find('a:ext',NS).get('cx'));height=int(pos.find('a:ext',NS).get('cy'))
                    if not (0.36*12192000<=width<=0.39*12192000 and height>=0.99*6858000):findings.append({'page':i,'problem':'Contents blue panel no longer matches the reference proportions'})
                    fill=panel.find('p:spPr/a:solidFill/a:srgbClr',NS)
                    if fill is None or fill.get('val','').upper()!='005BA6':findings.append({'page':i,'problem':'Contents panel must use reference blue'})
                if not any(n.startswith('ContentsRow') for n in shapes):findings.append({'page':i,'problem':'Missing chapter bars'})
            elif role=='summary':
                require('ReferenceHeader')
                labels={n[len('SummaryLabel'):] for n in shapes if n.startswith('SummaryLabel')}
                bodies={n[len('SummaryBody'):] for n in shapes if n.startswith('SummaryBody')}
                if not labels or labels!=bodies:findings.append({'page':i,'problem':'Summary topic blocks and explanation rows must correspond'})
            elif role=='cover':require('CoverBlueCurve')
            elif role=='closing':require('ClosingOuter');require('ClosingInner')
            elif role=='body-figure':require('ReferenceHeader')
            else:findings.append({'page':i,'problem':'Unknown layout_id'})
    return {'slides':len(roles),'findings':findings,'visual_match_not_proven':True}


if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('source');p.add_argument('--roles',nargs='+');p.add_argument('--outline');a=p.parse_args()
    roles=a.roles or [q['layout_id'] for q in json.loads(Path(a.outline).read_text(encoding='utf-8-sig'))['pages']]
    result=audit(a.source,roles);print(json.dumps(result,ensure_ascii=False));raise SystemExit(bool(result['findings']))
