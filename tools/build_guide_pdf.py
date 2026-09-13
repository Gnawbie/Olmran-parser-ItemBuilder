# -*- coding: utf-8 -*-
"""Regenerates USER_GUIDE.pdf directly from docs/guide.html, so the two
never drift apart and the PDF never needs its own hand-maintained copy
of every bullet point again. Builds a tiny DOM (via html.parser) of just
the content that matters (h1/tagline, h2/h3 headings, p, ul/ol/li,
including one level of list nesting), converts each element's inline
markup (b/i/code) into ReportLab's own supported Paragraph markup, and
walks that tree in document order to build the PDF's story.

Usage (run from the repo root, after editing docs/guide.html):
    python tools/build_guide_pdf.py docs/guide.html USER_GUIDE.pdf
    copy the result to both USER_GUIDE.pdf (repo root) and docs/USER_GUIDE.pdf

Kept in the repo (not a scratchpad-only script) specifically so it
survives between sessions - the previous hand-maintained generator was
lost when its scratchpad copy disappeared."""
import sys
import re
from html.parser import HTMLParser
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, ListFlowable, ListItem, HRFlowable
)

GUIDE_HTML_PATH = sys.argv[1] if len(sys.argv) > 1 else "guide.html"
OUT_PATH = sys.argv[2] if len(sys.argv) > 2 else "USER_GUIDE.pdf"

BLOCK_TAGS = {'h1', 'h2', 'h3', 'p', 'ul', 'ol', 'li'}
INLINE_TAGS = {'b': 'b', 'i': 'i', 'code': 'font'}


class Node:
    __slots__ = ('tag', 'attrs', 'children')
    def __init__(self, tag, attrs=None):
        self.tag = tag
        self.attrs = attrs or {}
        self.children = []  # list of Node or str


class GuideTreeBuilder(HTMLParser):
    """Builds a tiny tree of just h1/h2/h3/p/ul/ol/li (+ inline b/i/code/
    br inside them) - everything else (style, script, the topbar, the
    TOC card, footer) is skipped entirely by only ever pushing certain
    tags onto the stack and dropping text seen outside of one."""
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node('root')
        self.stack = [self.root]
        self.skip_depth = 0  # inside <style>/<script>/<div class="topbar">/<div class="toc">/footer
        self.skip_tag_stack = []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if self.skip_depth:
            if tag not in ('br',):
                self.skip_tag_stack.append(tag)
                self.skip_depth += 1
            return
        if tag in ('style', 'script', 'footer'):
            self.skip_tag_stack.append(tag)
            self.skip_depth += 1
            return
        if tag == 'div' and attrs.get('class') in ('topbar', 'toc'):
            self.skip_tag_stack.append(tag)
            self.skip_depth += 1
            return
        if tag in BLOCK_TAGS:
            node = Node(tag, attrs)
            self.stack[-1].children.append(node)
            self.stack.append(node)
        elif tag in INLINE_TAGS:
            node = Node(tag, attrs)
            self.stack[-1].children.append(node)
            self.stack.append(node)
        elif tag == 'br':
            self.stack[-1].children.append(Node('br'))
        # div/a/span etc. at this level are transparent - their text still
        # lands directly in whatever block/inline node is currently open.

    def handle_endtag(self, tag):
        if self.skip_depth:
            if self.skip_tag_stack and self.skip_tag_stack[-1] == tag:
                self.skip_tag_stack.pop()
                self.skip_depth -= 1
            return
        if tag in BLOCK_TAGS or tag in INLINE_TAGS:
            if len(self.stack) > 1 and self.stack[-1].tag == tag:
                self.stack.pop()

    def handle_data(self, data):
        if self.skip_depth:
            return
        self.stack[-1].children.append(data)


def render_inline(node_or_children):
    """Renders a block node's children (mixed str/Node) into one string
    using ReportLab's own Paragraph mini-markup (b/i supported directly;
    code becomes a monospace font tag) - safe to feed straight into
    Paragraph(...)."""
    def esc(s):
        return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

    def render_children(children):
        out = []
        for child in children:
            if isinstance(child, str):
                out.append(esc(child))
            elif child.tag == 'br':
                out.append('<br/>')
            elif child.tag in INLINE_TAGS:
                inner = render_children(child.children)
                if child.tag == 'code':
                    out.append(f'<font face="Courier" size="9">{inner}</font>')
                else:
                    out.append(f'<{child.tag}>{inner}</{child.tag}>')
            else:
                out.append(render_children(child.children))
        return ''.join(out)

    children = node_or_children.children if isinstance(node_or_children, Node) else node_or_children
    text = render_children(children)
    return re.sub(r'\s+', ' ', text).strip()


styles = getSampleStyleSheet()
title_style = ParagraphStyle('TitleX', parent=styles['Title'], fontSize=26, spaceAfter=4)
subtitle_style = ParagraphStyle('SubtitleX', parent=styles['Normal'], fontSize=12,
                                 textColor=colors.HexColor('#555555'), spaceAfter=18)
h1_style = ParagraphStyle('H1X', parent=styles['Heading1'], fontSize=18, spaceBefore=22, spaceAfter=8,
                          textColor=colors.HexColor('#1c2028'))
h2_style = ParagraphStyle('H2X', parent=styles['Heading2'], fontSize=13.5, spaceBefore=14, spaceAfter=6,
                          textColor=colors.HexColor('#2c3440'))
body_style = ParagraphStyle('BodyX', parent=styles['Normal'], fontSize=10, leading=14.5, spaceAfter=6,
                            alignment=TA_LEFT)
bullet_style = ParagraphStyle('BulletX', parent=body_style, leftIndent=0, spaceAfter=4)
toc_style = ParagraphStyle('TocX', parent=body_style, fontSize=10.5, leading=16)

story = []


def emit_list(ul_node, level=0):
    items = []
    for li in ul_node.children:
        if not isinstance(li, Node) or li.tag != 'li':
            continue
        # A nested <ul>/<ol> inside this <li> becomes its own indented
        # ListFlowable right after this item's own text.
        own_text_children = [c for c in li.children if not (isinstance(c, Node) and c.tag in ('ul', 'ol'))]
        nested_lists = [c for c in li.children if isinstance(c, Node) and c.tag in ('ul', 'ol')]
        text = render_inline(own_text_children)
        content = [Paragraph(text, bullet_style)]
        for nested in nested_lists:
            content.append(emit_list(nested, level + 1))
        if len(content) == 1:
            items.append(ListItem(content[0], leftIndent=14 + level * 14))
        else:
            from reportlab.platypus import KeepTogether
            items.append(ListItem(content, leftIndent=14 + level * 14))
    return ListFlowable(items, bulletType='bullet', start='•', leftIndent=10 + level * 4, spaceAfter=8)


def build_story(root):
    started = False
    for node in root.children:
        if not isinstance(node, Node):
            continue
        if not started:
            if node.tag == 'h2':
                started = True
            else:
                # Everything before the first <h2> (the h1 title and the
                # tagline <p>) is only ever used for the title page above -
                # never also dumped again as plain body text here.
                continue
        if node.tag == 'h2':
            story.append(Paragraph(render_inline(node), h1_style))
        elif node.tag == 'h3':
            story.append(Paragraph(render_inline(node), h2_style))
        elif node.tag == 'p':
            story.append(Paragraph(render_inline(node), body_style))
        elif node.tag in ('ul', 'ol'):
            story.append(emit_list(node))


def find_first(node, tag):
    for c in node.children:
        if isinstance(c, Node) and c.tag == tag:
            return c
    return None


def main():
    with open(GUIDE_HTML_PATH, encoding='utf-8') as f:
        html = f.read()
    parser = GuideTreeBuilder()
    parser.feed(html)
    root = parser.root

    h1 = find_first(root, 'h1')
    tagline_p = None
    for c in root.children:
        if isinstance(c, Node) and c.tag == 'p':
            tagline_p = c
            break
    title_text = render_inline(h1) if h1 else "Olmran Item Builder"
    tagline_text = render_inline(tagline_p) if tagline_p else ""
    version = tagline_text.split(' - ')[-1] if ' - ' in tagline_text else ''

    # ---------- Title page ----------
    story.append(Spacer(1, 60))
    story.append(Paragraph(title_text, title_style))
    story.append(Paragraph(f"Full Usage Guide - {version}" if version else tagline_text, subtitle_style))
    story.append(Paragraph(
        "A tool for parsing Olmran/MUD-style game logs, extracting loot data, building a master "
        "equipment database, and finding the optimal gear set for your character.", body_style))
    story.append(Spacer(1, 20))
    story.append(Paragraph(
        "This guide covers every tab and feature currently in the program. It's a companion to the "
        "README on GitHub, written for reading start-to-finish or jumping to the section you need.",
        body_style))
    story.append(PageBreak())

    # ---------- Table of contents (derived from every h2's own text) ----------
    story.append(Paragraph("Contents", h1_style))
    h2_nodes = [c for c in root.children if isinstance(c, Node) and c.tag == 'h2']
    for h2 in h2_nodes:
        story.append(Paragraph(render_inline(h2), toc_style))
    story.append(PageBreak())

    # ---------- Body: every h2/h3/p/ul/ol in original document order ----------
    build_story(root)

    doc = SimpleDocTemplate(OUT_PATH, pagesize=letter,
                            topMargin=54, bottomMargin=54, leftMargin=58, rightMargin=58,
                            title="Olmran Item Builder - Full Usage Guide")
    doc.build(story)
    print(f"Wrote {OUT_PATH} from {GUIDE_HTML_PATH} ({len(h2_nodes)} sections)")


if __name__ == '__main__':
    main()
