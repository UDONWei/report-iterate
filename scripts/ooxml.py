"""Read DOCX structure and perform narrowly scoped, structure-preserving edits.

Requires lxml. Word's display cache is inspected, not recomputed.
"""

from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from hashlib import sha256
import json
from pathlib import Path
import os
import re
from tempfile import NamedTemporaryFile
from zipfile import ZipFile

from lxml import etree as ET


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
DOC = "word/document.xml"
STORY = re.compile(r"word/(?:header[^/]*|footer[^/]*|footnotes|endnotes|comments)\.xml$")
CAPTION = re.compile(r"^(图|表|Figure|Table)\s*(\d+)(?:\s*[-–—.．]\s*(\d+))?", re.I)
HEAD_NAME = re.compile(r"(?:heading\s*|标题\s*)([1-9])$", re.I)
FIELD_CMD = re.compile(r'^\s*([a-z]+)(?:\s+(?:"([^"]+)"|([^\s\\]+)))?', re.I)
FIELD_ERROR = re.compile(
    r"Error!\s*(?:Reference source not found|Bookmark not defined)"
    r"|错误[！!]\s*(?:未找到引用源|找不到引用源|未定义书签)", re.I
)
LOCAL_LINK = re.compile(r'(?i)\\l\s+(?:"([^"]+)"|([^\s\\]+))')
CAP_TYPES = {"图": "图", "figure": "图", "表": "表", "table": "表"}
CONTAINERS = {W+n for n in ("sdt", "ins", "del", "moveFrom", "moveTo", "txbxContent",
                            "hyperlink", "smartTag", "customXml", "dir", "bdo")}
BARRIERS = {W+n for n in (
    "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd",
    "commentReference", "permStart", "permEnd", "footnoteReference",
    "endnoteReference", "fldChar", "drawing", "pict", "object", "sym",
    "lastRenderedPageBreak", "noBreakHyphen", "softHyphen",
    "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
)}
REF_HEADINGS = {"参考文献", "参考资料", "引用资料", "references", "bibliography", "works cited"}


class UnsafeEdit(ValueError):
    """The proposed edit is not an unambiguous, structurally safe replacement."""


def sha256_file(path):
    h = sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1024*1024), b""):
            h.update(b)
    return h.hexdigest()


def _digest(data):
    return sha256(json.dumps(data, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()


def _parse(data):
    root = ET.fromstring(data, ET.XMLParser(resolve_entities=False, no_network=True))
    if root.getroottree().docinfo.doctype:
        raise ValueError("DOCX XML with a document type declaration is unsupported")
    return root


def _val(element, path):
    p = element.find(path)
    return p.get(W+"val", "") if p is not None else ""


def _part_location(node):
    """Namespace prefix independent and 1-based, including same-tag siblings."""
    if node is None:
        return None
    p = []
    e = node
    while e is not None:
        name = ET.QName(e).localname
        if e.tag == W+"body":
            p.append("body")
            break
        if e.getparent() is None:
            p.append(name)
            break
        p.append(f"{name}[{1+sum(1 for _ in e.itersiblings(e.tag, preceding=True))}]")
        e = e.getparent()
    return "/".join(reversed(p))


@dataclass
class Paragraph:
    index: int
    location: str
    para_id: str | None
    section: tuple[str, ...]
    kind: str
    style: str
    text: str
    editable: bool
    outline_level: int | None = None


@dataclass
class _Span:
    start: int
    end: int
    node: object
    editable: bool


@dataclass
class _Buffer:
    pieces: list = field(default_factory=list)
    offsets: list = field(default_factory=list)
    barriers: set = field(default_factory=set)
    size: int = 0
    toc: bool = False

    def emit(self, s, node=None, editable=False):
        if s:
            self.offsets.append(_Span(self.size, self.size+len(s), node, editable))
            self.pieces.append(s)
            self.size += len(s)

    @property
    def text(self):
        return "".join(self.pieces)


@dataclass
class _Field:
    story: str
    start: str | None
    end: str | None = None
    encoding: str = "complex"
    instruction: str = ""
    result: str = ""
    complete: bool = False
    separated: bool = False

    def command(self):
        m = FIELD_CMD.match(self.instruction)
        return (m[1].upper(), m[2] or m[3] or "") if m else ("", "")


class Document:
    def __init__(self, path):
        self.path = Path(path).resolve(strict=True)
        with ZipFile(self.path) as z:
            names = z.namelist()
            if len(names) != len(set(names)) or DOC not in names:
                raise ValueError("ambiguous or missing DOCX package entry")
            self.source_signed = any(n.lower().startswith("_xmlsignatures/") for n in names)
            self.root = _parse(z.read(DOC))
            self.other = {n: _parse(z.read(n)) for n in names if STORY.fullmatch(n)}
            self.styles = _parse(z.read("word/styles.xml")) if "word/styles.xml" in names else None
            self._part_hashes = {n: sha256(z.read(n)).hexdigest() for n in names if n != DOC}
        self.body = self.root.find(W+"body")
        if self.body is None:
            raise ValueError("unsupported Word document namespace or empty body")
        self._source_hash = sha256_file(self.path)
        self._styles = {}
        self._default_style = ""
        if self.styles is not None:
            for s in self.styles.findall(W+"style"):
                ident = s.get(W+"styleId", "")
                self._styles[ident] = (_val(s, W+"name"), _val(s, W+"basedOn"),
                                      _val(s, f'{W}pPr/{W}outlineLvl'))
                if s.get(W+"type") == "paragraph" and s.get(W+"default") in ("true", "1"):
                    self._default_style = ident
        self._refresh()
        self._baseline_structure = self._structure().copy()

    def _style_chain(self, style):
        seen = set()
        while style and style not in seen:
            seen.add(style)
            name, inherited, level = self._styles.get(style, ("", "", ""))
            yield style, name, level
            style = inherited

    @staticmethod
    def _heading_from_val(value):
        if not value or not value.isdecimal():
            return None
        n = int(value)
        return n+1 if 0 <= n < 9 else None

    def _heading(self, node, chain):
        outline = node.find(f'{W}pPr/{W}outlineLvl')
        if outline is not None:
            return self._heading_from_val(outline.get(W+"val", ""))
        for style, name, outline in chain:
            if outline:
                return self._heading_from_val(outline)
            for option in (name, style):
                if m := HEAD_NAME.fullmatch(option):
                    return int(m[1])
        return None

    @staticmethod
    def _inside(node, tags):
        return any(a.tag in tags for a in node.iterancestors())

    def _walk(self, root, story):
        stacks, paras, active_bmarks = [], [], {}
        locs = {}

        def loc(p):
            if p is None:
                return None
            if p not in locs:
                locs[p] = _part_location(p)
            return locs[p]

        def barrier(b):
            if b is not None:
                b.barriers.add(b.size)

        def issue(code, p, **extra):
            self._issues.append({"severity": "error", "code": code, "story": story,
                                 "location": loc(p), **extra})

        walker = ET.iterwalk(root, events=("start", "end"))
        for event, e in walker:
            tag = e.tag
            if not isinstance(tag, str):
                continue
            if event == "start" and (tag in {W+"del", W+"moveFrom"} or (
                tag == MC+"Fallback" and e.getparent().find(MC+"Choice") is not None
            ) or (tag == MC+"Choice" and any(e.itersiblings(MC+"Choice", preceding=True)))):
                walker.skip_subtree()
                continue
            if tag == W+"p":
                if event == "start":
                    paras.append(e)
                    b = _Buffer()
                    self._buffers[e] = b
                    if story == DOC:
                        self._main_nodes.append(e)
                    if any(f.command()[0] == "TOC" for f in stacks):
                        b.toc = True
                else:
                    if paras and paras[-1] is e:
                        paras.pop()
                continue
            p = paras[-1] if paras else None
            b = self._buffers[p] if p is not None else None
            if tag == W+"fldSimple":
                barrier(b)
                if event == "start":
                    f = _Field(story, loc(p), encoding="simple",
                               instruction=e.get(W+"instr", ""), separated=True)
                    self._fields.append(f)
                    stacks.append(f)
                    if b and f.command()[0] == "TOC":
                        b.toc = True
                elif stacks and stacks[-1].encoding == "simple":
                    f = stacks.pop()
                    f.end = loc(p)
                    f.complete = True
                else:
                    issue("unmatched_simple_field", p)
                continue
            if event != "start":
                continue
            if tag in BARRIERS:
                barrier(b)
            if tag == W+"fldChar":
                mode = e.get(W+"fldCharType")
                if mode == "begin":
                    f = _Field(story, loc(p))
                    self._fields.append(f)
                    stacks.append(f)
                elif stacks and stacks[-1].encoding == "complex":
                    f = stacks[-1]
                    if mode == "separate":
                        if f.separated:
                            issue("multiple_field_separators", p)
                        f.separated = True
                        if b and f.command()[0] == "TOC":
                            b.toc = True
                    elif mode == "end":
                        stacks.pop()
                        f.end = loc(p)
                        f.complete = True
                        if b and f.command()[0] == "TOC":
                            b.toc = True
                    else:
                        issue("unknown_field_character", p, field_character=mode)
                else:
                    issue("unmatched_field_character", p, field_character=mode)
            elif tag == W+"instrText":
                if stacks and stacks[-1].encoding == "complex" and not stacks[-1].separated:
                    stacks[-1].instruction += e.text or ""
                else:
                    issue("orphan_field_instruction", p)
            elif tag == W+"bookmarkStart":
                ident, name = e.get(W+"id"), e.get(W+"name")
                if ident in active_bmarks or not ident or not name:
                    issue("invalid_bookmark_start", p, id=ident, name=name)
                record = {"story": story, "name": name, "id": ident, "start": loc(p), "end": None,
                          "start_element": _part_location(e), "end_element": None}
                self._bookmarks.append(record)
                active_bmarks[ident] = record
            elif tag == W+"bookmarkEnd":
                ident = e.get(W+"id")
                if ident in active_bmarks:
                    record = active_bmarks.pop(ident)
                    record["end"] = loc(p)
                    record["end_element"] = _part_location(e)
                else:
                    issue("unmatched_bookmark_end", p, id=ident)
            elif tag == W+"hyperlink":
                if name := e.get(W+"anchor"):
                    self._links.append({"story": story, "location": loc(p), "kind": "hyperlink",
                                        "target": name, "result": None})
            elif tag in {W+"t", W+"tab", W+"br", W+"cr", W+"noBreakHyphen", W+"softHyphen"}:
                if b is None or self._inside(e, {W+"pPr", W+"rPr"}):
                    continue
                s = {W+"tab": "\t", W+"br": "\n", W+"cr": "\n",
                     W+"noBreakHyphen": "\u2011", W+"softHyphen": "\u00ad"}.get(tag, e.text or "")
                if stacks and not all(f.separated for f in stacks):
                    continue
                for f in stacks:
                    f.result += s
                parent = e.getparent()
                allowed = (
                    tag == W+"t" and not stacks and not active_bmarks
                    and parent.tag == W+"r" and parent.getparent() is p
                    and parent.find(f'{W}rPr/{W}vanish') is None
                    and not self._inside(e, CONTAINERS)
                )
                b.emit(s, e, allowed)
        for f in stacks:
            self._issues.append({"severity": "error", "code": "unterminated_field", "story": story,
                                 "location": f.start, "instruction": f.instruction})
        for record in active_bmarks.values():
            self._issues.append({"severity": "error", "code": "unterminated_bookmark", "story": story,
                                 "location": record["start"], "name": record["name"]})

    def _refresh(self):
        self._buffers, self._main_nodes, self._fields = {}, [], []
        self._issues, self._bookmarks, self._links = [], [], []
        self._walk(self.body, DOC)
        for story, root in sorted(self.other.items()):
            self._walk(root, story)
        by_location = defaultdict(list)
        for f in self._fields:
            cmd, arg = f.command()
            if f.story == DOC:
                by_location[f.start].append(f)
            if (cmd in {"REF", "PAGEREF", "NOTEREF"} and arg
                    or cmd == "HYPERLINK" and LOCAL_LINK.search(f.instruction)):
                if cmd == "HYPERLINK":
                    m = LOCAL_LINK.search(f.instruction)
                    arg = m[1] or m[2]
                self._links.append({"story": f.story, "location": f.start, "kind": cmd,
                                    "target": arg, "result": f.result})
            if FIELD_ERROR.search(f.result):
                self._issues.append({"severity": "error", "code": "cached_field_error", "story": f.story,
                                     "location": f.start, "result": f.result})
        names = defaultdict(list)
        for b in self._bookmarks:
            if b["name"] and b["end_element"]:
                names[b["name"]].append(b)
        for link in self._links:
            target = names.get(link["target"], [])
            if not target:
                state = "missing"
            elif len(target) > 1:
                state = "ambiguous"
            else:
                state = "resolved"
            link.update(status=state, destination=target[0]["start"] if len(target) == 1 else None,
                        target_story=target[0]["story"] if len(target) == 1 else None)
            if state != "resolved":
                self._issues.append({"severity": "error", "code": f"{state}_reference_target",
                                     "story": link["story"], "location": link["location"],
                                     "target": link["target"]})
        self.paragraphs, self._caption_rows = [], []
        section = []
        for i, p in enumerate(self._main_nodes):
            b = self._buffers[p]
            location = _part_location(p)
            text = b.text
            style = _val(p, f'{W}pPr/{W}pStyle') or self._default_style
            chain = list(self._style_chain(style))
            style_keys = {v.lower() for st, name, _ in chain for v in (st, name)}
            native = [
                (f, CAP_TYPES[a.lower()]) for f in by_location[location]
                for cmd, a in [f.command()]
                if cmd == "SEQ" and a.lower() in CAP_TYPES and f.complete
            ]
            table = bool(next(p.iterancestors(W+"tbl"), None) is not None)
            container = self._inside(p, CONTAINERS)
            toc = b.toc or any(re.fullmatch(r"toc\s*\d+", s, flags=re.I) for s in style_keys)
            level = None if toc or native or table or container else self._heading(p, chain)
            keytext = text.strip().casefold()
            if not (toc or table or container or native) and keytext in REF_HEADINGS:
                # A fair number of reports use a separate unnumbered bibliography heading.
                level = level or 1
            if native:
                kind = "native_caption"
                for f, caption_type in native:
                    m = CAPTION.match(text.lstrip())
                    label = None
                    title = text.strip()
                    if m and CAP_TYPES[m[1].lower()] == caption_type:
                        label = caption_type + str(int(m[2])) + (f"-{int(m[3])}" if m[3] else "")
                        title = text.lstrip()[m.end():].strip(" :：\t")
                    else:
                        self._issues.append({"severity": "hint", "code": "unresolved_caption_label",
                                             "story": DOC, "location": location, "text": text})
                    self._caption_rows.append({"index": i, "location": location, "kind": kind,
                                              "type": caption_type, "label": label, "title": title,
                                              "text": text, "instruction": f.instruction,
                                              "section": list(t for _, t in section)})
            elif toc:
                kind = "toc"
            elif style_keys.intersection({"caption", "figurecaption", "tablecaption", "题注", "图题", "表题"}):
                stripped = text.lstrip()
                m = CAPTION.match(stripped)
                tail = stripped[m.end():] if m else ""
                if m and (not tail or tail[0].isspace() or tail[0] in ":："):
                    kind = "manual_caption"
                    caption_type = CAP_TYPES[m[1].lower()]
                    label = caption_type + str(int(m[2])) + (f"-{int(m[3])}" if m[3] else "")
                    self._caption_rows.append({"index": i, "location": location, "kind": kind,
                                              "type": caption_type, "label": label, "title": tail.strip(" :：\t"),
                                              "text": text, "instruction": None,
                                              "section": list(t for _, t in section)})
                else:
                    kind = "caption_style_candidate"
                    if stripped:
                        self._issues.append({"severity": "hint", "code": "caption_style_without_label",
                                             "story": DOC, "location": location, "text": text})
            elif level is not None:
                kind = "heading"
                section = [(l, t) for l, t in section if l < level] + [(level, text.strip())]
            elif container:
                kind = "container"
            elif any(t.casefold() in REF_HEADINGS for _, t in section):
                kind = "references"
            elif table:
                kind = "table"
            else:
                kind = "body"
            self.paragraphs.append(Paragraph(
                i, location, p.get(W14+"paraId"), tuple(t for _, t in section),
                kind, style, text, kind in {"body", "table"} and any(t.editable for t in b.offsets), level
            ))
        previous = {}
        native_labels = defaultdict(list)
        for c in self._caption_rows:
            if c["kind"] != "native_caption" or not c["label"]:
                continue
            native_labels[c["label"]].append(c["location"])
            nums = tuple(int(s) for s in re.findall(r"\d+", c["label"]))
            if (p := previous.get(c["type"])) and len(p[0]) == len(nums) and p[0] > nums:
                self._issues.append({"severity": "hint", "code": "caption_order", "story": DOC,
                                     "location": c["location"], "label": c["label"],
                                     "previous_location": p[1]})
            previous[c["type"]] = (nums, c["location"])
        for label, locations in native_labels.items():
            if len(locations) > 1:
                self._issues.append({"severity": "error", "code": "duplicate_native_caption_label",
                                     "story": DOC, "location": locations[0], "label": label,
                                     "locations": locations})

    def _fields_data(self):
        return [dict(asdict(f), command=f.command()[0], argument=f.command()[1]) for f in self._fields]

    def _table_data(self):
        indices = {n: p.index for n, p in zip(self._main_nodes, self.paragraphs)}
        out = []
        for table in self.body.iter(W+"tbl"):
            rows = []
            for row in table.iter(W+"tr"):
                if next(row.iterancestors(W+"tbl"), None) is not table:
                    continue
                cells = []
                for cell in row.iter(W+"tc"):
                    if next(cell.iterancestors(W+"tr"), None) is not row:
                        continue
                    ps = [
                        indices[p] for p in cell.iter(W+"p") if p in indices
                        and next(p.iterancestors(W+"tc"), None) is cell
                    ]
                    cells.append({"paragraph_indices": ps,
                                  "text": "\n".join(self.paragraphs[p].text for p in ps),
                                  "grid_span": _val(cell, f'{W}tcPr/{W}gridSpan') or "1"})
                rows.append(cells)
            out.append({"location": _part_location(table), "rows": rows})
        return out

    def _structure(self):
        # Replace only structurally eligible w:t in the canonical fingerprints.
        mutable = {span.node for p, node in zip(self.paragraphs, self._main_nodes) if p.editable
                   for span in self._buffers[node].offsets if span.editable}
        # Empty w:t can be created by a replacement across equally formatted runs.
        eligible = {p.index for p in self.paragraphs if p.kind in {"body", "table"}}
        for i in eligible:
            p = self._main_nodes[i]
            for r in p.findall(W+"r"):
                for t in r.findall(W+"t"):
                    if not t.text:
                        # Keep empty text and XML-space changes out of the comparison;
                        # field result and bookmark texts are hashed additionally.
                        mutable.add(t)
        backup = [(n, n.text, n.get(XML_SPACE)) for n in mutable]
        try:
            for n, _, _ in backup:
                n.text = None
                n.attrib.pop(XML_SPACE, None)
            skeleton = sha256(ET.tostring(self.root, method="c14n", with_comments=True)).hexdigest()
        finally:
            for n, content, space in backup:
                n.text = content
                if space is not None:
                    n.set(XML_SPACE, space)
        return {"fields_sha256": _digest(self._fields_data()),
                "bookmarks_sha256": _digest(self._bookmarks),
                "captions_sha256": _digest(self._caption_rows),
                "body_skeleton_sha256": skeleton,
                "unchanged_parts_sha256": _digest(self._part_hashes)}

    def inspect(self, include_paragraphs=True):
        out = {
            "file": str(self.path), "source_sha256": self._source_hash,
            "structure": self._structure(), "paragraph_count": len(self.paragraphs),
            "headings": [asdict(p) for p in self.paragraphs if p.kind == "heading"],
            "captions": self._caption_rows, "fields": self._fields_data(),
            "bookmarks": self._bookmarks, "refs": self._links, "issues": self._issues,
            "table_count": sum(1 for _ in self.body.iter(W+"tbl")),
        }
        if include_paragraphs:
            out["paragraphs"] = [asdict(p) for p in self.paragraphs]
            out["tables"] = self._table_data()
        return out

    def replace(self, index, old, new):
        if not isinstance(index, int) or not 0 <= index < len(self.paragraphs):
            raise UnsafeEdit("invalid paragraph index")
        if not isinstance(old, str) or not old or not isinstance(new, str) or old == new:
            raise UnsafeEdit("replacement requires non-empty old text and genuinely different new text")
        if any(ord(c) < 32 or ord(c) in {0xfffe, 0xffff} or 0xd800 <= ord(c) <= 0xdfff
               for c in old+new):
            raise UnsafeEdit("line breaks, tabs or invalid XML text require explicit editing")
        if self.source_signed:
            raise UnsafeEdit("editing a digitally signed package requires a separate native workflow")
        p = self.paragraphs[index]
        if not p.editable:
            raise UnsafeEdit(f"protected paragraph: {p.kind} at {p.location}")
        start = p.text.find(old)
        if start < 0 or p.text.find(old, start+1) >= 0:
            raise UnsafeEdit(f"replacement must match exactly once at {p.location}")
        end = start+len(old)
        b = self._buffers[self._main_nodes[index]]
        selected = [s for s in b.offsets if s.start < end and s.end > start]
        if not selected or not all(s.editable for s in selected) or any(start < u < end for u in b.barriers):
            raise UnsafeEdit("replacement crosses or edits a field, bookmark, link or protected content")

        def run_style(span):
            rpr = span.node.getparent().find(W+"rPr")
            return ET.tostring(rpr, method="c14n") if rpr is not None else None

        if len({run_style(s) for s in selected}) > 1:
            raise UnsafeEdit("replacement spans differently formatted runs")
        # A newly added literal may not masquerade as a live citation or a numbered-reference field.
        tokens = re.compile(r"\[\d+\]|(?:图|表|Figure\s+|Table\s+)\s*\d+[-–—.]\d+", re.I)
        if Counter(tokens.findall(old)) != Counter(tokens.findall(new)):
            raise UnsafeEdit("editing a literal figure/table reference or citation requires separate verification")
        backup = [(s.node, s.node.text, s.node.get(XML_SPACE)) for s in selected]
        try:
            for i, s in enumerate(selected):
                before = s.node.text or ""
                rstart, rend = max(0, start-s.start), min(end-s.start, len(before))
                after = before[:rstart] + (new if i == 0 else "") + before[rend:]
                s.node.text = after
                if after and (after[0].isspace() or after[-1].isspace()):
                    s.node.set(XML_SPACE, "preserve")
            self._refresh()
            if self.paragraphs[index].text != p.text[:start] + new + p.text[end:]:
                raise UnsafeEdit("replacement did not produce the expected display text")
            if self._structure() != self._baseline_structure:
                raise UnsafeEdit("replacement changes protected document structure")
        except Exception:
            for t, content, space in backup:
                t.text = content
                if space is None:
                    t.attrib.pop(XML_SPACE, None)
                else:
                    t.set(XML_SPACE, space)
            self._refresh()
            raise
        return self.paragraphs[index]

    def save(self, path):
        p = Path(path)
        if (p.resolve() == self.path or p.exists() and p.samefile(self.path)):
            raise UnsafeEdit("do not overwrite or alias the input DOCX")
        if sha256_file(self.path) != self._source_hash:
            raise UnsafeEdit("the input DOCX changed since extraction")
        if self._structure() != self._baseline_structure:
            raise UnsafeEdit("protected document structure changed")
        p.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(prefix=".report-iterate-", suffix=".tmp", dir=p.parent, delete=False) as f:
            tmp = Path(f.name)
        try:
            with ZipFile(self.path) as src, ZipFile(tmp, "w") as dst:
                dst.comment = src.comment
                for entry in src.infolist():
                    data = (ET.tostring(self.root, encoding="UTF-8", xml_declaration=True, standalone=True)
                            if entry.filename == DOC else src.read(entry))
                    dst.writestr(entry, data)
            probe = Document(tmp)
            if probe._baseline_structure != self._baseline_structure:
                raise UnsafeEdit("saved DOCX failed the structural comparison")
            if [p.text for p in probe.paragraphs] != [p.text for p in self.paragraphs]:
                raise UnsafeEdit("saved DOCX failed the text comparison")
            result = probe._source_hash
            os.replace(tmp, p)
            return result
        finally:
            tmp.unlink(missing_ok=True)
