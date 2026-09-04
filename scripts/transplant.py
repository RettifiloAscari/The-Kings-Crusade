#!/usr/bin/env python3
"""Transplant QS document content into the DnD 5e Style Template package."""
import re, shutil, zipfile, os, sys

# --- self-bootstrapping template location ---
# Looks for an already-extracted template dir; if absent, extracts it from the
# .docx sitting alongside this script (or in the project/uploads directories).
_HERE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(_HERE, "tpl", "DnD_5e_StyleTemplate")

def _decode_template(md_path):
    """Rebuild the template .docx from a base64-encoded markdown carrier."""
    import base64
    txt = open(md_path, encoding="utf-8").read()
    m = re.search(r"```base64\s*(.*?)```", txt, re.S)
    if not m:
        return None
    raw = base64.b64decode("".join(m.group(1).split()))
    tmp = os.path.join(_HERE, "_template_decoded.docx")
    open(tmp, "wb").write(raw)
    return tmp

if not os.path.exists(os.path.join(TPL, "word", "document.xml")):
    # Prefer a real .docx if one is present (chat upload retains binaries);
    # otherwise decode the base64 carrier, which is what survives project storage.
    _docx = [
        os.path.join(_HERE, "DnD_5e_StyleTemplate.docx"),
        "/mnt/user-data/uploads/DnD_5e_StyleTemplate.docx",
    ]
    _b64 = [
        os.path.join(_HERE, "style_template_encoded.md"),
        "/mnt/project/style_template_encoded.md",
        "/mnt/user-data/uploads/style_template_encoded.md",
    ]
    _src = None
    for p in _docx:
        if os.path.exists(p):
            try:
                zipfile.ZipFile(p).namelist()   # must be a real zip, not text
                _src = p
                break
            except Exception:
                pass
    if _src is None:
        for p in _b64:
            if os.path.exists(p):
                _src = _decode_template(p)
                if _src:
                    break
    if _src is None:
        raise SystemExit(
            "Template not found. Provide style_template_encoded.md (project "
            "knowledge) or a genuine DnD_5e_StyleTemplate.docx (chat upload)."
        )
    os.makedirs(TPL, exist_ok=True)
    with zipfile.ZipFile(_src) as _z:
        _z.extractall(TPL)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# --- read template plumbing ---
tpl_doc = open(f"{TPL}/word/document.xml", encoding="utf-8").read()
tpl_rels = open(f"{TPL}/word/_rels/document.xml.rels", encoding="utf-8").read()

# footer1 rId
m = re.search(r'Id="(rId\d+)"[^>]*Target="footer1\.xml"', tpl_rels)
FOOTER_RID = m.group(1)

# template margins from its second (content) sectPr
sects = re.findall(r"<w:sectPr.*?</w:sectPr>", tpl_doc, re.S)
mg = re.search(r"<w:pgMar[^>]*/>", sects[1] if len(sects) > 1 else sects[0])
PGMAR = mg.group(0)
PGSZ = '<w:pgSz w:h="16838" w:w="11906" w:orient="portrait"/>'

def sect_xml(cols, with_footer):
    foot = (f'<w:footerReference w:type="default" r:id="{FOOTER_RID}"/>'
            f'<w:footerReference w:type="even" r:id="{FOOTER_RID}"/>') if with_footer else ""
    c = f'<w:cols w:num="{cols}" w:space="360"/>' if cols > 1 else '<w:cols w:num="1"/>'
    return foot + PGSZ + PGMAR + c

IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

def merge_images(doc, srcdir, work):
    """Carry images from the generated .docx into the template package.

    The template is a complete .docx with its own media and its own relationship
    ids, and this script previously copied only document.xml and numbering.xml
    across. Any image would therefore reference an rId that does not exist in the
    template's rels, and its bytes would never be copied -- so the picture is
    silently dropped or the document fails to load.

    Three things have to happen, in this order:
      1. copy word/media/* from the source package into the template package;
      2. re-map the source image relationship ids, because they collide with the
         template's (docx-js starts at rId1, and the template uses rId1-rId16
         for its own styles, fonts, headers, footers and hyperlinks);
      3. declare the file extension in [Content_Types].xml if it is not already.

    docx-js names media files by content hash, so the names cannot collide with
    the template's image1.png / image2.png, and identical images de-duplicate.
    Returns the rewritten document.xml text.
    """
    src_rels_path = os.path.join(srcdir, "word", "_rels", "document.xml.rels")
    if not os.path.exists(src_rels_path):
        return doc
    src_rels = open(src_rels_path, encoding="utf-8").read()
    imgs = re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*/>', src_rels)
    imgs = [(rid, tgt) for rid, tgt in imgs
            if re.search(r'Id="%s"[^>]*Type="%s"' % (re.escape(rid), re.escape(IMAGE_REL)), src_rels)]
    if not imgs:
        return doc

    work_rels_path = os.path.join(work, "word", "_rels", "document.xml.rels")
    work_rels = open(work_rels_path, encoding="utf-8").read()
    used = [int(n) for n in re.findall(r'Id="rId(\d+)"', work_rels)]
    next_id = (max(used) if used else 0) + 1

    media_dir = os.path.join(work, "word", "media")
    os.makedirs(media_dir, exist_ok=True)

    remap, exts, added = {}, set(), []
    for rid, target in imgs:
        src_file = os.path.join(srcdir, "word", target.replace("/", os.sep))
        if not os.path.exists(src_file):
            continue
        shutil.copy(src_file, os.path.join(work, "word", target.replace("/", os.sep)))
        new_id = "rId%d" % next_id
        next_id += 1
        remap[rid] = new_id
        exts.add(os.path.splitext(target)[1].lstrip(".").lower())
        added.append('<Relationship Id="%s" Type="%s" Target="%s"/>' % (new_id, IMAGE_REL, target))

    if not remap:
        return doc

    work_rels = work_rels.replace("</Relationships>", "".join(added) + "</Relationships>")
    open(work_rels_path, "w", encoding="utf-8").write(work_rels)

    # Remap in one pass: replacing ids one at a time can chain (rId7 -> rId17,
    # then a later rule rewrites that rId17 again).
    doc = re.sub(r'r:embed="([^"]+)"',
                 lambda m: 'r:embed="%s"' % remap.get(m.group(1), m.group(1)), doc)

    ct_path = os.path.join(work, "[Content_Types].xml")
    ct = open(ct_path, encoding="utf-8").read()
    for ext in sorted(exts):
        if 'Extension="%s"' % ext not in ct:
            mime = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png",
                    "gif": "gif", "bmp": "bmp"}.get(ext, ext)
            ct = ct.replace("</Types>",
                            '<Default ContentType="image/%s" Extension="%s"/></Types>' % (mime, ext))
    open(ct_path, "w", encoding="utf-8").write(ct)
    return doc


# Generators mark a table as full-width by giving it style FULLWIDTH_STYLE. The
# style is never defined anywhere -- it is only a marker this script looks for and
# then strips, so nothing downstream sees a dangling style reference.
FULLWIDTH_STYLE = "KCFullWidth"
FULLWIDTH_MARKER = '<w:tblStyle w:val="%s"/>' % FULLWIDTH_STYLE

def wrap_fullwidth_tables(doc, two_col):
    """Let marked tables span both columns of a two-column body.

    A three-column table with a prose column is unreadable at half the page width;
    it wraps to three or four words a line. OOXML has no "span the columns" flag on
    a table, so the fix is a pair of continuous section breaks around it: the
    paragraph before the table closes the two-column section, and the paragraph
    after it closes the one-column section that the table now sits in. Text resumes
    in two columns afterwards, from the body sectPr.

    Both marker paragraphs are empty and set to an exact one-twip line so they add
    no visible space. The type must be explicitly continuous -- the OOXML default is
    nextPage, which would throw every full-width table onto a page of its own.
    """
    def _wrap(m):
        blk = m.group(0)
        if FULLWIDTH_MARKER not in blk:
            return blk
        blk = blk.replace(FULLWIDTH_MARKER, "")
        if not two_col:
            return blk          # already full width; nothing to do but drop the marker
        def brk(cols):
            return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/>'
                    '<w:sectPr><w:type w:val="continuous"/>'
                    + sect_xml(cols, True) + '</w:sectPr></w:pPr></w:p>')
        return brk(2) + blk + brk(1)

    return re.sub(r"<w:tbl>.*?</w:tbl>", _wrap, doc, flags=re.S)


def carry_missing_styles(doc, srcdir, work):
    """Carry over any paragraph style the content uses and the template lacks.

    The transplant deliberately keeps the template's styles.xml, because that is
    where the visual identity lives -- the template's Heading1 must beat the
    generator's. But docx-js also emits styles of its own that the template never
    heard of, and the most important is ListParagraph: every bullet carries
    <w:pStyle w:val="ListParagraph"/>, and with no such style defined the
    paragraph falls back to the template's default indent and LibreOffice drops
    the bullet entirely. Bullets rendered as plain indented prose, in every
    document, invisibly -- the numbering was intact the whole time.

    So: template wins wherever both define a style; anything only the source
    defines is carried across.
    """
    src_styles_path = os.path.join(srcdir, "word", "styles.xml")
    work_styles_path = os.path.join(work, "word", "styles.xml")
    if not (os.path.exists(src_styles_path) and os.path.exists(work_styles_path)):
        return
    src_styles = open(src_styles_path, encoding="utf-8").read()
    work_styles = open(work_styles_path, encoding="utf-8").read()

    have = set(re.findall(r'w:styleId="([^"]+)"', work_styles))
    used = set(re.findall(r'<w:pStyle w:val="([^"]+)"/>', doc))

    carried = []
    for style_id in sorted(used - have):
        m = re.search(r'<w:style [^>]*w:styleId="%s".*?</w:style>' % re.escape(style_id),
                      src_styles, re.S)
        if m:
            carried.append(m.group(0))
    if not carried:
        return
    work_styles = work_styles.replace("</w:styles>", "".join(carried) + "</w:styles>")
    open(work_styles_path, "w", encoding="utf-8").write(work_styles)


# --- breathing room under a table ---
# A table carries no "space after" in OOXML, and the body paragraphs are authored with
# spacing.after only, so prose immediately following a table butts against its bottom
# border with no gap at all. Headings already look right, because their style supplies
# spacing.before -- which is why the defect only ever showed up on table-then-prose and
# never on table-then-heading.
#
# Fixing it here rather than in the generators is deliberate: table() returns a single
# Table object, so a per-call fix would mean changing every call site in every generator
# (forty-four in one repository, seventy-one in the other) and remembering it forever
# after. One pass over the assembled XML covers every table, including any added later.
TABLE_GAP = 180   # twips; a shade under the 200 that separates two body paragraphs

_PARA_AFTER_TBL = re.compile(r'(</w:tbl>\s*)(<w:p\b(?:(?!</w:p>).)*?</w:p>)', re.S)
_IS_HEADING = re.compile(r'w:pStyle w:val="(Heading\d|Title|Subtitle)"')
_SPACING = re.compile(r'<w:spacing\b([^/>]*)/>')
_TEXT = re.compile(r'<w:t[^>]*>([^<]*)</w:t>')

def gap_after_tables(doc):
    """Give the first paragraph after each table a space-before."""
    def fix(m):
        head, para = m.group(1), m.group(2)
        if _IS_HEADING.search(para):
            return m.group(0)          # its style already supplies the gap
        if 'w:before=' in para:
            return m.group(0)          # author asked for something specific; respect it
        if not ''.join(_TEXT.findall(para)).strip():
            # A blank paragraph IS already the gap. The stat-block helper pushes one after
            # its ability table, and it emits <w:t></w:t> rather than no run at all, so the
            # test has to be on the text content and not on the tag.
            return m.group(0)
        s = _SPACING.search(para)
        if s:
            # add the attribute to the spacing element that is already there
            para = para[:s.start()] + '<w:spacing w:before="%d"%s/>' % (TABLE_GAP, s.group(1)) + para[s.end():]
        else:
            open_tag = re.match(r'<w:p\b[^>]*>', para)
            if '<w:pPr>' in para:
                para = para.replace('<w:pPr>', '<w:pPr><w:spacing w:before="%d"/>' % TABLE_GAP, 1)
            else:
                para = (para[:open_tag.end()]
                        + '<w:pPr><w:spacing w:before="%d"/></w:pPr>' % TABLE_GAP
                        + para[open_tag.end():])
        return head + para
    return _PARA_AFTER_TBL.sub(fix, doc)

# --- a heading must not be left at the foot of a column ---
# The template's Heading styles carry keepNext, which stops a heading being the very
# last line in a column -- but keepNext binds one block, and one block is often not
# enough. Two shapes kept slipping through:
#
#   * "Tiered Skill DCs", a two-line lead-in, and then the table that IS the section,
#     which went over the break on its own;
#   * a gazetteer entry -- "Cairn Ithel" plus its one-line italic stat rule -- sitting
#     at the foot of the last column on a page with every word of the place on the next.
#
# Both read as a title announcing nothing. The fix is two more bindings, applied here
# for the same reason gap_after_tables is: table() returns a single Table and headings
# are helpers, so a per-call fix would mean touching every call site in eighteen
# generators and remembering it forever after.
#
#   A. A paragraph immediately followed by a table keeps with it, so a lead-in can
#      never be separated from the table it introduces.
#   B. A short paragraph immediately after a heading keeps with what follows, so the
#      heading drags a stat rule or a one-line preamble AND the body under it.
#
# Both are cheap. keepNext binds the last line of a block to the first line of the
# next, so a long paragraph still splits normally and only its tail travels; the
# whitespace cost is bounded by the short paragraph in rule B.
SHORT_LEAD_2COL = 110   # characters; about two lines in the 3.28in column
SHORT_LEAD_1COL = 220   # ... and about two lines across the single-column page

_KEEP_NEXT = '<w:keepNext w:val="1"/>'
_BLOCK = re.compile(r'<w:p\b(?:(?!</w:p>).)*?</w:p>|<w:p\b[^>]*/>|\x00TBL\d+\x00', re.S)
_PPR_HEAD = re.compile(r'<w:pPr>(<w:pStyle [^>]*/>)?')


def _with_keep_next(para):
    """Add keepNext to a paragraph, in the one place the schema allows it."""
    if 'w:keepNext' in para:
        return para
    m = _PPR_HEAD.search(para)
    if m:
        # CT_PPr is a sequence: pStyle, then keepNext, then everything else.
        return para[:m.end()] + _KEEP_NEXT + para[m.end():]
    m = re.match(r'<w:p\b[^>]*>', para)
    if not m:
        return para
    return para[:m.end()] + '<w:pPr>' + _KEEP_NEXT + '</w:pPr>' + para[m.end():]


def bind_headings(doc, two_col=True):
    """Keep a heading with enough of its section to be worth reading."""
    short = SHORT_LEAD_2COL if two_col else SHORT_LEAD_1COL

    # Mask the tables first. Their cells are full of paragraphs, and a keepNext on a
    # paragraph inside a cell binds nothing and confuses the block walk.
    tbls = []

    def stash(m):
        tbls.append(m.group(0))
        return "\x00TBL%d\x00" % (len(tbls) - 1)

    masked = re.sub(r'<w:tbl>.*?</w:tbl>', stash, doc, flags=re.S)

    blocks = list(_BLOCK.finditer(masked))
    edits = []
    for i, b in enumerate(blocks):
        cur = b.group(0)
        if cur.startswith('\x00') or _IS_HEADING.search(cur):
            continue                     # a table, or a heading whose style already binds
        if not ''.join(_TEXT.findall(cur)).strip():
            continue                     # a spacer paragraph binds nothing worth keeping
        nxt = blocks[i + 1].group(0) if i + 1 < len(blocks) else ''
        prev = blocks[i - 1].group(0) if i else ''
        lead_in_to_table = nxt.startswith('\x00')
        opens_a_section = (_IS_HEADING.search(prev)
                           and len(''.join(_TEXT.findall(cur)).strip()) <= short)
        if lead_in_to_table or opens_a_section:
            edits.append(b)

    for b in reversed(edits):            # from the end, so earlier spans stay valid
        masked = masked[:b.start()] + _with_keep_next(b.group(0)) + masked[b.end():]

    return re.sub(r'\x00TBL(\d+)\x00', lambda m: tbls[int(m.group(1))], masked)


def convert(src, dst, two_col=True):
    work = os.path.join(_HERE, "work_tpl")
    shutil.rmtree(work, ignore_errors=True)
    shutil.copytree(TPL, work)

    # bring in OUR content + numbering
    srcdir = os.path.join(_HERE, "work_src")
    shutil.rmtree(srcdir, ignore_errors=True)
    os.makedirs(srcdir)
    with zipfile.ZipFile(src) as z:
        z.extractall(srcdir)
    doc = open(f"{srcdir}/word/document.xml", encoding="utf-8").read()
    shutil.copy(f"{srcdir}/word/numbering.xml", f"{work}/word/numbering.xml")
    doc = merge_images(doc, srcdir, work)
    carry_missing_styles(doc, srcdir, work)

    # --- section surgery on our document.xml ---
    # replace final sectPr with template-derived one (continuous so body starts on the title page)
    cont = '<w:type w:val="continuous"/>' if two_col else ''
    doc = re.sub(r"<w:sectPr.*?</w:sectPr>",
                 "<w:sectPr>" + cont + sect_xml(2 if two_col else 1, True) + "</w:sectPr>",
                 doc, count=1, flags=re.S)

    # restyle title block: first paragraph -> Title style, second -> Subtitle
    paras = list(re.finditer(r"<w:p\b.*?</w:p>", doc, re.S))
    def restyle(pblock, style):
        blk = pblock
        blk = re.sub(r"<w:sz w:val=\"\d+\"/>", "", blk)
        blk = re.sub(r"<w:szCs w:val=\"\d+\"/>", "", blk)
        blk = re.sub(r"<w:b/>", "", blk)
        if "<w:pPr>" in blk:
            blk = blk.replace("<w:pPr>", '<w:pPr><w:pStyle w:val="%s"/>' % style, 1)
        else:
            blk = blk.replace(">", '><w:pPr><w:pStyle w:val="%s"/></w:pPr>' % style, 1)
        return blk
    if len(paras) >= 2:
        p0, p1 = paras[0].group(0), paras[1].group(0)
        doc = doc.replace(p0, restyle(p0, "Title"), 1)
        doc = doc.replace(p1, restyle(p1, "Subtitle"), 1)

    # ability-table cells: sz 20 -> 17 so headers fit two-column width
    doc = gap_after_tables(doc)
    doc = bind_headings(doc, two_col)

    doc = doc.replace('<w:sz w:val="20"/>', '<w:sz w:val="17"/>')

    if two_col:
        # insert a single-column section break before the FIRST Heading1 paragraph
        h1 = re.search(r'<w:p\b[^>]*>(?:(?!</w:p>).)*?w:val="Heading1"', doc, re.S)
        if h1:
            brk = ('<w:p><w:pPr><w:sectPr>' + sect_xml(1, True) + '</w:sectPr></w:pPr></w:p>')
            doc = doc[:h1.start()] + brk + doc[h1.start():]

    doc = wrap_fullwidth_tables(doc, two_col)

    open(f"{work}/word/document.xml", "w", encoding="utf-8").write(doc)

    # --- repackage ---
    if os.path.exists(dst):
        os.remove(dst)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(work):
            for f in files:
                full = os.path.join(root, f)
                arc = os.path.relpath(full, work)
                z.write(full, arc)
    print("converted:", os.path.basename(dst))

if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    two_col = "--single" not in sys.argv
    convert(src, dst, two_col)
