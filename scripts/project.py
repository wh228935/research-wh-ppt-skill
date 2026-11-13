"""Portable project approvals, phase checks and immutable PPT revision bundles.

Only standard library. Approval quotes must be supplied from real user replies;
this program checks file identity, never the authenticity of human consent.
"""
import argparse
import hashlib
import json
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def sha(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def package_hash():
    h = hashlib.sha256()
    for p in sorted(ROOT.rglob('*')):
        rel = p.relative_to(ROOT)
        if not p.is_file() or any(x in {'.git', '__pycache__', 'work', 'local',
                                        'outputs', 'validation-private'} for x in rel.parts):
            continue
        if p.suffix in {'.pyc', '.log'}:
            continue
        h.update(rel.as_posix().encode('utf-8'))
        h.update(bytes.fromhex(sha(p)))
    return h.hexdigest()


def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def save(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + '.tmp')
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    temp.replace(path)


def locate(work, value):
    p = Path(value)
    return (p if p.is_absolute() else Path(work) / p).resolve()


def now():
    return datetime.now(timezone.utc).isoformat()


def load_state(work):
    return read(Path(work) / 'project.json')


def validate_snapshot(path, seen=None):
    """A review JSON can bind child files as [{path,sha256}], recursively."""
    seen = seen or set()
    p = Path(path).resolve()
    if p in seen:
        return
    seen.add(p)
    if p.suffix.lower() != '.json':
        return
    obj = read(p)
    if not isinstance(obj, dict):
        return
    for item in obj.get('files', []):
        child = locate(p.parent, item['path'])
        if not child.is_file() or sha(child) != item['sha256']:
            raise ValueError(f'审核附件已变更或缺失: {child}')
        validate_snapshot(child, seen)
    if 'source_sha256' in obj and isinstance(obj.get('slides'), list):
        for item in obj['slides']:
            child = locate(p.parent, item['file'])
            if not child.is_file() or sha(child) != item['sha256']:
                raise ValueError(f'审核预览已变更或缺失: {child}')


def approved(state, scope, target=None):
    r = state.get('approvals', {}).get(scope)
    if not r:
        raise ValueError(f'缺少真实用户确认: {scope}')
    p = Path(r['path'])
    if target and p.resolve() != Path(target).resolve():
        raise ValueError(f'{scope}确认的文件与当前计划不一致')
    if not p.is_file() or sha(p) != r['sha256']:
        raise ValueError(f'{scope}确认已失效，文件改变或缺失: {p}')
    if not r.get('quote', '').strip():
        raise ValueError(f'{scope}未记录真实回复')
    validate_snapshot(p)
    return p


def norm(text):
    return re.sub(r'\s+', ' ', str(text)).strip()


def object_text(obj):
    if obj['type'] == 'text':
        return [obj.get('text', '')]
    if obj['type'] == 'table':
        return [str(v) for row in obj['rows'] for v in row]
    return []


def check_plan(state, work, plan_path, stage):
    if package_hash() != state['skill']['sha256']:
        raise ValueError('当前skill与任务固定版本不同：使用原版本，或先提出迁移方案由用户决定')
    intake_path = approved(state, 'intake')
    intake = read(intake_path)
    if not isinstance(intake.get('duration_minutes'), (int, float)) or intake['duration_minutes'] <= 0:
        raise ValueError('intake未确认正数汇报时长')
    if state['mode'] == 'edit':
        for scope in ['edit-scope', 'edit-reading', 'edit-plan']:
            approved(state, scope)
        if stage not in {'edit', 'deliver'}:
            raise ValueError('局改模式不使用新建生成器，使用原稿编辑工具并保持原风格')
        edit_path = approved(state, 'edit-plan')
        edit_plan = read(edit_path)
        for permission in edit_plan.get('permissions', []):
            approved(state, permission['scope'], locate(edit_path.parent, permission['proposal']))
        for asset in edit_plan.get('assets', []):
            required = list(asset.get('required_permissions', []))
            kind = asset.get('origin', 'source')
            if kind=='external': required.append('add-image:'+asset['id'])
            if kind=='processed': required.append('transform:'+asset['id'])
            if kind=='redrawn': required.append('redraw:'+asset['id'])
            for scope in required:
                if not any(p['scope']==scope for p in edit_plan.get('permissions', [])):
                    raise ValueError(f'局改缺少图件专项提案: {scope}')
        return []
    docs = intake.get('documents', [])
    if not docs:
        raise ValueError('未登记本次正文阅读范围')
    for doc in docs:
        approved(state, f"reading-overview:{doc['id']}")
        if not doc.get('chapters'):
            raise ValueError(f"{doc['id']}缺少正文各章清单")
        source = locate(intake_path.parent, doc['path'])
        if not source.is_file() or sha(source) != doc['sha256']:
            raise ValueError(f'原文已改变或缺失，需复核受影响阅读: {source}')
        for chapter in doc['chapters']:
            approved(state, f"reading-chapter:{doc['id']}:{chapter['id']}")
    figures_path = approved(state, 'figures')
    figures = read(figures_path)
    figure_files = {locate(figures_path.parent, x['path']): x['sha256'] for x in figures.get('files', [])}
    approved(state, 'content')
    plan_path = Path(plan_path).resolve()
    plan = read(plan_path)
    style = read(locate(plan_path.parent, plan['style_file']))
    if style.get('font') not in {'Microsoft YaHei', '微软雅黑'} and not any(x['scope']=='font-change' for x in plan.get('permissions', [])):
        raise ValueError('字体替换需要font-change具体提案和用户确认')
    outline_path = locate(plan_path.parent, plan['outline_file'])
    approved(state, 'outline', outline_path)
    outline = read(outline_path)
    outline_pages = {p['id']: p for p in outline['pages']}
    pages = plan['pages']
    if len({p['id'] for p in pages}) != len(pages):
        raise ValueError('页面ID重复')
    selection_path = approved(state, 'prototype-selection')
    chosen = read(selection_path)['pages']
    if not 3 <= len(chosen) <= 5 or len(set(chosen)) != len(chosen):
        raise ValueError('代表页清单必须为3—5张不同页面')
    if any(x not in outline_pages for x in chosen):
        raise ValueError('代表页不在确认的大纲里')
    if stage == 'prototype':
        if set(p['id'] for p in pages) != set(chosen):
            raise ValueError('首轮代表页必须与确认清单一致，不能夹带全稿或变体')
    elif stage in {'full', 'deliver'}:
        prototype_review = read(approved(state, 'prototypes'))
        style_path = locate(plan_path.parent, plan['style_file'])
        bound = {locate(Path(state['approvals']['prototypes']['path']).parent, x['path']):
                 x['sha256'] for x in prototype_review.get('files', [])}
        if style_path not in bound or bound[style_path] != sha(style_path):
            raise ValueError('代表页确认没有绑定当前风格参数文件')
        if set(p['id'] for p in pages) != set(outline_pages):
            raise ValueError('全稿页面与确认大纲不一致，不能自行增删')
    assets = {a['id']: a for a in plan.get('assets', [])}
    selected = set(figures.get('selected_assets', []))
    for permission in plan.get('permissions', []):
        approved(state, permission['scope'], locate(plan_path.parent, permission['proposal']))
    for page in pages:
        if page['id'] not in outline_pages:
            raise ValueError(f"页面不在确认大纲: {page['id']}")
        op = outline_pages[page['id']]
        if norm(page['title']) != norm(op['title']):
            raise ValueError(f"标题不符: {page['id']}")
        allowed_text = {norm(x) for x in op.get('on_screen', [])}
        allowed_text.update({norm(page['title']), norm(op.get('header', ''))})
        if norm(page.get('header', '')) != norm(op.get('header', '')):
            raise ValueError('章节页眉改变，先更新大纲确认')
        if not str(page.get('notes', '')).strip():
            raise ValueError(f"缺少详细备注: {page['id']}")
        for obj in page.get('objects', []):
            for text in object_text(obj):
                if norm(text) and norm(text) not in allowed_text:
                    raise ValueError(f"{page['id']}含未经大纲确认的文字: {text[:60]}")
            if obj['type'] == 'image':
                aid = obj['asset_id']
                if aid not in selected or aid not in op.get('assets', []):
                    raise ValueError(f'图件未选定或不在当页大纲: {aid}')
                asset = assets[aid]
                ap = locate(plan_path.parent, asset['path'])
                if sha(ap) != asset['sha256']:
                    raise ValueError(f'图件文件已改变: {aid}')
                origin = asset.get('origin', 'source')
                required = list(asset.get('required_permissions', []))
                if origin not in {'source','external','processed','redrawn'}:
                    raise ValueError('未知图件来源类别')
                if origin in {'source','external'}:
                    if figure_files.get(ap) != asset['sha256']:
                        raise ValueError(f'选图确认未绑定此图的实际文件版本: {aid}')
                else:
                    original = locate(plan_path.parent, asset['original_path'])
                    if figure_files.get(original) != asset['original_sha256'] or sha(original) != asset['original_sha256']:
                        raise ValueError(f'处理/重绘的原始图件未绑定选图确认: {aid}')
                if origin=='external': required.append('add-image:'+aid)
                if origin=='processed': required.append('transform:'+aid)
                if origin=='redrawn': required.append('redraw:'+aid)
                for scope in required:
                    matches = [x for x in plan.get('permissions', []) if x['scope'] == scope]
                    if not matches:
                        raise ValueError(f'缺少图件专项提案: {scope}')
                if obj.get('crop') or obj.get('fit', 'contain') != 'contain':
                    raise ValueError('生成器只等比容纳完整图片，裁剪先经批准并另生成处理文件')
    return [p['id'] for p in pages]


def archive(state, work, manifest_path, output, revision):
    if not re.fullmatch(r'v\d{2,}', revision):
        raise ValueError('稿次格式应为v01、v02等')
    bundle = read(manifest_path)
    dest = Path(output).resolve() / revision
    if dest.exists():
        raise ValueError(f'稿次已存在，拒绝覆盖: {dest}')
    from office import audit_pptx, docx_paragraphs, ppt_slides
    base = Path(manifest_path).parent
    ppt = locate(base, bundle['pptx'])
    word = locate(base, bundle['docx'])
    report = audit_pptx(ppt)
    if report['errors']:
        raise ValueError('PPT结构检查未通过: ' + '; '.join(report['errors']))
    slides = ppt_slides(ppt)
    text = '\n'.join(docx_paragraphs(word))
    for i, slide in enumerate(slides, 1):
        if f'第{i}页' not in text or slide['title'] not in text:
            raise ValueError('Word页序/标题与最终PPT不对应')
        if norm(slide['notes']) not in norm(text):
            raise ValueError('Word讲稿与最终PPT备注不对应')
    source_hash = sha(ppt)
    preview_dir = locate(base, bundle['preview_dir'])
    previews = read(preview_dir / 'render-report.json')
    if previews['source_sha256'] != source_hash or len(previews['slides']) != len(slides):
        raise ValueError('预览与最终PPT不是同一版')
    for entry in previews['slides']:
        if sha(preview_dir / entry['file']) != entry['sha256']:
            raise ValueError('预览文件已改变')
    thumb = locate(base, bundle['contact_sheet'])
    thumb_record = read(str(thumb) + '.source.json')
    if thumb_record.get('source_sha256') != source_hash or thumb_record.get('contact_sheet_sha256') != sha(thumb):
        raise ValueError('缩略图没有绑定最终PPT，或缩略图文件改变')
    if bundle['formal_pages'] + bundle.get('backup_pages', 0) != len(slides):
        raise ValueError('正式页与备份页数量不等于实际页数')
    qa_path = locate(base, bundle['qa'])
    qa = read(qa_path)
    if qa.get('source_sha256') != source_hash or qa.get('visual_reviewed') is not True:
        raise ValueError('QA未绑定最终PPT或尚未实际逐页目视检查')
    if not qa.get('content_reviewed') or qa.get('unresolved_errors'):
        raise ValueError('内容未检查或存在未解决错误')
    for name in ['PowerPoint', 'WPS']:
        q = locate(base, bundle['software_reports'][name])
        r = read(q)
        if r.get('application') != name or not r.get('opened') or r.get('source_sha256') != source_hash or len(r.get('slides', [])) != len(slides):
            raise ValueError(f'{name}实际软件报告缺失或不是最终文件')
    dest.parent.mkdir(parents=True, exist_ok=True)
    temp = dest.with_name(dest.name + '.staging')
    if temp.exists():
        raise ValueError(f'存在未处理归档暂存目录: {temp}')
    temp.mkdir()
    try:
        shutil.copy2(ppt, temp / ppt.name)
        shutil.copy2(word, temp / word.name)
        (temp / 'previews').mkdir()
        for entry in previews['slides']:
            shutil.copy2(preview_dir / entry['file'], temp / 'previews' / Path(entry['file']).name)
        shutil.copy2(thumb, temp / ('contact-sheet' + thumb.suffix))
        files = [{'path': p.relative_to(temp).as_posix(), 'sha256': sha(p)}
                 for p in sorted(temp.rglob('*')) if p.is_file()]
        save(temp / 'version-record.json', {'revision': revision, 'skill': state['skill'],
              'created_utc': now(), 'files': files, 'formal_pages': bundle['formal_pages'],
              'backup_pages': bundle.get('backup_pages', 0)})
        temp.rename(dest)
    except Exception:
        # Leave staging evidence for inspection; never recursively remove unknown files.
        raise
    return str(dest)


def main():
    p = argparse.ArgumentParser(description=__doc__)
    sub = p.add_subparsers(dest='cmd', required=True)
    a = sub.add_parser('init'); a.add_argument('--workdir', required=True)
    a.add_argument('--mode', choices=['new', 'edit'], default='new')
    for name in ['status', 'record', 'gate', 'archive']:
        a = sub.add_parser(name); a.add_argument('--workdir', required=True)
        if name == 'record':
            a.add_argument('--scope', required=True); a.add_argument('--file', required=True)
            a.add_argument('--quote', required=True)
        if name == 'gate':
            a.add_argument('--plan'); a.add_argument('--stage', choices=['prototype','full','edit','deliver'], required=True)
        if name == 'archive':
            a.add_argument('--manifest', required=True); a.add_argument('--output', required=True)
            a.add_argument('--revision', required=True); a.add_argument('--plan', required=True)
    a = sub.add_parser('snapshot'); a.add_argument('--output', required=True); a.add_argument('files', nargs='+')
    a = sub.add_parser('fingerprint')
    args = p.parse_args()
    if args.cmd == 'fingerprint':
        print(package_hash()); return
    if args.cmd == 'snapshot':
        output = Path(args.output).resolve()
        if output.exists():
            raise ValueError('审核快照已存在，请另存新文件')
        save(output, {'files': [{'path': str(Path(x).resolve()), 'sha256': sha(x)} for x in args.files]})
        print(output); return
    work = Path(args.workdir).resolve()
    if work == ROOT or ROOT in work.parents:
        raise ValueError('任务目录不能放在skill包中')
    if args.cmd == 'init':
        if (work / 'project.json').exists():
            raise ValueError('项目已存在，拒绝覆盖')
        save(work / 'project.json', {'mode': args.mode, 'created_utc': now(),
             'skill': {**read(ROOT / 'VERSION.json'), 'sha256': package_hash()}, 'approvals': {}})
        print(work / 'project.json'); return
    state = load_state(work)
    if args.cmd == 'record':
        if not args.quote.strip():
            raise ValueError('用户回复不能为空')
        target = locate(work, args.file)
        validate_snapshot(target)
        record = {'path': str(target), 'sha256': sha(target), 'quote': args.quote,
                  'recorded_utc': now()}
        old = state['approvals'].get(args.scope)
        if old:
            state.setdefault('approval_history', []).append({'scope': args.scope, **old})
        state['approvals'][args.scope] = record
        save(work / 'project.json', state)
        print(json.dumps({'recorded': args.scope, 'note': '记录工具不证明回复真实性'}, ensure_ascii=False))
    elif args.cmd == 'status':
        status = {}
        for scope in state['approvals']:
            try:
                approved(state, scope); status[scope] = 'file-unchanged'
            except (ValueError, OSError) as e:
                status[scope] = str(e)
        print(json.dumps({'mode': state['mode'], 'skill': state['skill'], 'approvals': status}, ensure_ascii=False, indent=2))
    elif args.cmd == 'gate':
        pages = check_plan(state, work, locate(work, args.plan) if args.plan else work / 'plan.json', args.stage)
        print(json.dumps({'stage': args.stage, 'pages': pages, 'file_checks_passed': True,
                          'human_truth_not_proven': True}, ensure_ascii=False))
    elif args.cmd == 'archive':
        check_plan(state, work, locate(work, args.plan), 'deliver')
        print(archive(state, work, locate(work, args.manifest), args.output, args.revision))


if __name__ == '__main__':
    try:
        main()
    except (ValueError, OSError, KeyError, TypeError) as exc:
        print(f'ERROR: {exc}', file=sys.stderr)
        sys.exit(2)
