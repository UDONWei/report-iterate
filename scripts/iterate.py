"""Run reviewable, resumable batches of text edits against immutable Word copies."""

from __future__ import annotations

import argparse
from contextlib import contextmanager
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import sys
from tempfile import NamedTemporaryFile, TemporaryDirectory

from ooxml import Document, UnsafeEdit, sha256_file


class BatchError(ValueError):
    pass


ID = re.compile(r'[A-Za-z0-9][A-Za-z0-9_.-]{0,79}\Z')
HASH = re.compile(r'[a-f0-9]{64}\Z')
PENDING = {'review', 'evidence_pending', 'blocked'}


def require(ok, message):
    if not ok:
        raise BatchError(message)


def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def write_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile('w', encoding='utf-8', newline='\n', dir=path.parent, delete=False) as f:
        tmp = Path(f.name)
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write('\n')
        f.flush()
        os.fsync(f.fileno())
    try:
        os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)


@contextmanager
def exclusive(work):
    """An OS file lock is released even if the process is interrupted."""
    with (work / 'run.lock').open('a+b') as lock:
        lock.seek(0)
        try:
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(lock.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            raise BatchError('run directory is locked by another processor') from exc
        try:
            yield
        finally:
            lock.seek(0)
            if os.name == 'nt':
                msvcrt.locking(lock.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(lock.fileno(), fcntl.LOCK_UN)


def plan_info(path):
    path = Path(path).resolve(strict=True)
    raw = path.read_bytes()
    plan = json.loads(raw.decode('utf-8-sig'))
    require(isinstance(plan, dict) and plan.get('format') == 1, 'expected plan format 1')
    require(isinstance(plan.get('documents'), list) and plan['documents'], 'documents must be a nonempty list')
    require(isinstance(plan.get('items'), list) and plan['items'], 'items must be a nonempty list')
    seen, names, docs = set(), set(), {}
    for d in plan['documents']:
        idx = d['id']
        source = (path.parent / d['path']).resolve()
        require(ID.fullmatch(idx) and idx not in docs, f'invalid or duplicate document id: {idx}')
        require(HASH.fullmatch(d['sha256']), f'missing or invalid baseline SHA256: {idx}')
        require(source.suffix.lower() == '.docx' and source.name.casefold() not in names, f'expected distinct .docx file names: {idx}')
        names.add(source.name.casefold())
        docs[idx] = (source, d['sha256'])
    for item in plan['items']:
        idx = item['id']
        require(ID.fullmatch(idx) and idx not in seen, f'invalid or duplicate item id: {idx}')
        seen.add(idx)
        require(item.get('risk') in {'ordinary', 'major'}, f'invalid risk: {idx}')
        require(item.get('evidence') in {'verified', 'not_required', 'pending'}, f'invalid evidence status: {idx}')
        for key in ('title', 'rationale'):
            require(isinstance(item.get(key), str) and item[key].strip(), f'{key} required: {idx}')
        require(item['evidence'] != 'verified' or isinstance(item.get('sources'), list) and item['sources'], f'verified evidence needs source pointers: {idx}')
        changes = item['changes']
        require(isinstance(changes, list) and changes, f'missing change locations: {idx}')
        if item['risk'] == 'major':
            require(isinstance(item.get('impact'), str) and item['impact'].strip(), f'impact required: {idx}')
            require(isinstance(item.get('options'), list) and len(item['options']) >= 2, f'at least two options required: {idx}')
        spans = set()
        for ch in changes:
            a = ch['anchor']
            require(ch['document'] in docs, f'unknown document: {idx}')
            require(isinstance(a['location'], str) and a['location'] and isinstance(a['section'], list), f'invalid location or section: {idx}')
            require(all(isinstance(a[k], str) for k in ('paragraph', 'before', 'after')), f'invalid anchor text or neighbors: {idx}')
            require(all(isinstance(ch[k], str) for k in ('old', 'new')) and ch['old'] and ch['new'] != ch['old'], f'invalid or no-op replacement: {idx}')
            span = (ch['document'], a['location'])
            require(span not in spans, f'combine multiple changes to one paragraph in a single replacement: {idx}')
            spans.add(span)
    return plan, docs, hashlib.sha256(raw).hexdigest()


def immutable(work, digest):
    require(HASH.fullmatch(digest), 'invalid snapshot hash')
    return work / 'objects' / (digest + '.docx')


def put(work, source, digest):
    dst = immutable(work, digest)
    if not dst.exists():
        with NamedTemporaryFile(dir=dst.parent, delete=False) as f:
            tmp = Path(f.name)
            with Path(source).open('rb') as src:
                shutil.copyfileobj(src, f)
            f.flush()
            os.fsync(f.fileno())
        try:
            require(sha256_file(tmp) == digest, f'snapshot copy did not match the expected hash: {Path(source).name}')
            os.replace(tmp, dst)
        finally:
            tmp.unlink(missing_ok=True)
    require(sha256_file(dst) == digest, f'corrupt snapshot: {dst.name}')
    return dst


def state_for(work, documents, batch_hash):
    path = work / 'state.json'
    if path.exists():
        state = read_json(path)
        require(state.get('format') == 1 and state['plan_sha256'] == batch_hash,
                'work order changed; use a new run directory and the verified exports as your new baseline')
        require(set(state['documents']) == set(documents), 'snapshot manifest has different documents')
        for idx, (src, digest) in documents.items():
            entry = state['documents'][idx]
            require(entry['baseline'] == digest and entry['file'] == src.name, 'baseline differs from earlier manifest')
            for h in {entry['baseline'], entry['current']}:
                require(sha256_file(immutable(work, h)) == h, f'damaged snapshot for {idx}')
        return state
    (work / 'objects').mkdir(exist_ok=True)
    state = {'format': 1, 'plan_sha256': batch_hash, 'documents': {}, 'items': {}}
    for idx, (src, digest) in documents.items():
        put(work, src, digest)
        state['documents'][idx] = {'file': src.name, 'baseline': digest, 'current': digest}
    write_json(path, state)
    return state


def neighbors(doc, p):
    ps = doc.paragraphs
    return (ps[p.index - 1].text if p.index else '', ps[p.index + 1].text if p.index + 1 < len(ps) else '')


def anchor_for(doc, p):
    before, after = neighbors(doc, p)
    return {'location': p.location, 'section': list(p.section), 'paragraph': p.text,
            'before': before, 'after': after, 'para_id': p.para_id}


def locate(doc, anchor):
    hits = [p for p in doc.paragraphs if p.location == anchor['location'] and list(p.section) == anchor['section']]
    if anchor.get('para_id'):
        hits = [p for p in hits if p.para_id == anchor['para_id']]
    require(len(hits) == 1, f"location is not unique or no longer matches: {anchor['location']}")
    return hits[0]


def execute(item, work, state, hook):
    item_documents, before, receipts = {}, {}, []
    for ch in item['changes']:
        idx, a = ch['document'], ch['anchor']
        entry = state['documents'][idx]
        if idx not in item_documents:
            base = Document(immutable(work, entry['baseline']))
            current = Document(immutable(work, entry['current']))
            item_documents[idx] = base, current
            before[idx] = current.inspect(include_paragraphs=False)['structure']
        base, doc = item_documents[idx]
        b = locate(base, a)
        prev, nxt = neighbors(base, b)
        require(b.text == a['paragraph'] and (prev, nxt) == (a['before'], a['after']), f"baseline paragraph or neighbors mismatch: {idx} {a['location']}")
        old, new = ch['old'], ch['new']
        require(b.text.count(old) == 1, f'original fragment is absent or ambiguous: {idx} {a["location"]}')
        p = locate(doc, a)
        require(p.editable, f'paragraph is protected ({p.kind}): {idx} {a["location"]}')
        expect = b.text.replace(old, new, 1)
        if p.text == expect:
            status = 'already_present'
        else:
            require(p.text == b.text, f'paragraph has a different edit; rebase this item: {idx} {a["location"]}')
            ret = doc.replace(p.index, old, new)
            require(ret.text == expect, f'written paragraph differs from planned result: {idx}')
            status = 'applied'
        receipts.append({'document': idx, 'location': a['location'], 'status': status, 'old': old, 'new': new})

    changes = {}
    with TemporaryDirectory(dir=work) as scratch:
        for idx, (_, doc) in item_documents.items():
            if not any(q['document'] == idx and q['status'] == 'applied' for q in receipts):
                continue
            dst = Path(scratch) / (idx + '.docx')
            digest = doc.save(dst)
            require(Document(dst).inspect(include_paragraphs=False)['structure'] == before[idx], f'protected Word structure changed in {idx}')
            put(work, dst, digest)
            changes[idx] = digest
    if hook:
        hook(item, changes)  # test seam after snapshots exist, before atomic commit
    return changes, receipts


def run(path, work, limit=None, hook=None):
    plan, documents, plan_hash = plan_info(path)
    work = Path(work).resolve()
    require(limit is None or limit > 0, 'stop-after must be positive')
    require(isinstance(plan['items'], list), 'missing work orders')
    for src, _ in documents.values():
        require(src != work and not src.is_relative_to(work), 'baseline file must be outside the run directory')
    if work.exists() and not ((work / 'state.json').exists() or (work / 'run.lock').exists()):
        require(not any(work.iterdir()), 'use an empty or initialized run directory')
    work.mkdir(parents=True, exist_ok=True)
    changed, processed, reused = 0, 0, 0
    with exclusive(work):
        state = state_for(work, documents, plan_hash)
        require(set(state['items']) <= {x['id'] for x in plan['items']}, 'state contains unknown items')
        for item in plan['items']:
            idx = item['id']
            if idx in state['items']:
                reused += 1
                continue
            if limit is not None and processed >= limit:
                break
            if item['risk'] == 'major':
                rec = {'status': 'review', 'options': item['options'], 'impact': item['impact'], 'locations': item['changes']}
            elif item['evidence'] == 'pending':
                rec = {'status': 'evidence_pending', 'reason': item['rationale'], 'locations': item['changes']}
            else:
                try:
                    update, receipts = execute(item, work, state, hook)
                except (BatchError, UnsafeEdit) as exc:
                    rec = {'status': 'blocked', 'reason': str(exc), 'locations': item['changes']}
                else:
                    rec = {'status': 'applied' if update else 'already_present', 'changes': receipts,
                           'document_hashes': {d: {'before': state['documents'][d]['current'], 'after': h} for d, h in update.items()}}
                    for d, h in update.items():
                        state['documents'][d]['current'] = h
                    changed += sum(c['status'] == 'applied' for c in receipts)
            state['items'][idx] = rec
            write_json(work / 'state.json', state)
            processed += 1
        remaining = len(plan['items']) - len(state['items'])
        pending = {k: v['status'] for k, v in state['items'].items() if v['status'] in PENDING}
        summary = {'new_edits': changed, 'new_records': processed, 'reused_records': reused,
                   'not_processed': remaining, 'pending': pending, 'state': str(work / 'state.json')}
        print(json.dumps(summary, ensure_ascii=False))
        return 75 if remaining else (3 if pending else 0)


def export(work, output):
    work, output = Path(work).resolve(strict=True), Path(output).resolve()
    require(not output.exists() and not output.is_relative_to(work) and not work.is_relative_to(output),
            'export must use a new directory separate from run storage')
    with exclusive(work):
        state = read_json(work / 'state.json')
        files = []
        for d in state['documents'].values():
            src = immutable(work, d['current'])
            require(sha256_file(src) == d['current'], 'damaged snapshot')
            require(Path(d['file']).name == d['file'] and Path(d['file']).suffix.lower() == '.docx', 'invalid export filename')
            files.append((src, d['file']))
        output.mkdir(parents=True, exist_ok=False)
        for src, name in files:
            shutil.copyfile(src, output / name)
            require(sha256_file(output / name) == src.stem, 'export did not match committed snapshot')
        write_json(output / 'receipts.json', state)


def main():
    p = argparse.ArgumentParser(description=__doc__)
    s = p.add_subparsers(dest='command', required=True)
    inspect = s.add_parser('inspect')
    inspect.add_argument('docx', type=Path)
    inspect.add_argument('--out', type=Path)
    inspect.add_argument('--summary', action='store_true')
    compare = s.add_parser('compare')
    compare.add_argument('before', type=Path)
    compare.add_argument('after', type=Path)
    r = s.add_parser('run')
    r.add_argument('plan', type=Path)
    r.add_argument('work', type=Path)
    r.add_argument('--stop-after', type=int)
    e = s.add_parser('export')
    e.add_argument('work', type=Path)
    e.add_argument('output', type=Path)
    a = p.parse_args()
    try:
        if a.command == 'run':
            return run(a.plan, a.work, a.stop_after)
        if a.command == 'export':
            export(a.work, a.output)
            return 0
        if a.command == 'compare':
            first, second = [Document(x).inspect(include_paragraphs=False) for x in (a.before, a.after)]
            keys = set(first['structure']) | set(second['structure'])
            unequal = sorted(k for k in keys if first['structure'].get(k) != second['structure'].get(k))
            print(json.dumps({'same_protected_structure': not unequal, 'different': unequal,
                              'issues_before': first['issues'], 'issues_after': second['issues']}, ensure_ascii=False, indent=2))
            return int(bool(unequal))
        require(a.docx.suffix.lower() == '.docx', 'expected a .docx file')
        doc = Document(a.docx)
        out = doc.inspect(include_paragraphs=not a.summary)
        out['file_sha256'] = sha256_file(a.docx)
        if not a.summary:
            out['anchors'] = [anchor_for(doc, p) | {'kind': p.kind, 'editable': p.editable} for p in doc.paragraphs]
        if a.out:
            require(not a.out.exists() and a.out.suffix.lower() == '.json', 'inspect outputs must be new .json files')
            write_json(a.out, out)
        else:
            print(json.dumps(out, ensure_ascii=False, indent=2))
        return 0
    except (ValueError, OSError, KeyError, TypeError, UnsafeEdit) as exc:
        print(f'{type(exc).__name__}: {exc}', file=sys.stderr)
        return 2


if __name__ == '__main__':
    sys.exit(main())
