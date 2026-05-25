#!/usr/bin/env python3
"""Generate PPTX: Tracing Emotion Dynamics in Academic Writing"""

import io, numpy as np, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE as SHAPE

# ── PALETTE ───────────────────────────────────────────────────────────────────
ROSE       = RGBColor(0xe8, 0xa4, 0xb4)
ROSE_DARK  = RGBColor(0xc4, 0x76, 0x8a)
PEACH      = RGBColor(0xf5, 0xc8, 0xa0)
BLUSH      = RGBColor(0xfd, 0xf0, 0xf3)
WARM       = RGBColor(0xfd, 0xfb, 0xf9)
WHITE      = RGBColor(0xff, 0xff, 0xff)
DARK       = RGBColor(0x1e, 0x1e, 0x1e)
GRAY       = RGBColor(0x66, 0x66, 0x66)
LIGHT      = RGBColor(0xaa, 0xaa, 0xaa)
BORDER     = RGBColor(0xf0, 0xdd, 0xe3)
CARD       = RGBColor(0xfd, 0xfb, 0xf9)
FAINT_ROSE = RGBColor(0xf5, 0xdc, 0xe4)

FD = "Georgia"      # display / headers
FB = "Calibri"      # body
W, H = 13.33, 7.5  # slide dimensions in inches

prs = Presentation()
prs.slide_width  = Inches(W)
prs.slide_height = Inches(H)
blank = prs.slide_layouts[6]   # completely blank layout


# ── LOW-LEVEL HELPERS ─────────────────────────────────────────────────────────

def new_slide():
    return prs.slides.add_slide(blank)

def set_bg(sl, rgb):
    f = sl.background.fill; f.solid(); f.fore_color.rgb = rgb

def rct(sl, l, t, w, h, fill=WHITE, line_rgb=None, line_pt=0.5):
    s = sl.shapes.add_shape(SHAPE.RECTANGLE,
                            Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line_rgb:
        s.line.color.rgb = line_rgb; s.line.width = Pt(line_pt)
    else:
        s.line.fill.background()
    return s

def ovl(sl, l, t, w, h, fill=BLUSH):
    s = sl.shapes.add_shape(SHAPE.OVAL,
                            Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s

def _run(p, text, sz, bold, italic, col, font):
    r = p.add_run()
    r.text = text; r.font.name = font; r.font.size = Pt(sz)
    r.font.bold = bold; r.font.italic = italic; r.font.color.rgb = col

def txt(sl, text, l, t, w, h, sz=12, bold=False, italic=False,
        col=DARK, align=PP_ALIGN.LEFT, font=FB, wrap=True):
    tb = sl.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    _run(p, text, sz, bold, italic, col, font)
    return tb, tf

def add_para(tf, text, sz=12, bold=False, italic=False,
             col=DARK, align=PP_ALIGN.LEFT, font=FB, sb=0, sa=2):
    p = tf.add_paragraph(); p.alignment = align
    if sb: p.space_before = Pt(sb)
    if sa: p.space_after  = Pt(sa)
    _run(p, text, sz, bold, italic, col, font)
    return p

# ── SLIDE CHROME HELPERS ──────────────────────────────────────────────────────

def watermark(sl):
    txt(sl, 'corpus.bfsu.edu.cn', (W-2.6)/2, H-0.38, 2.6, 0.28,
        sz=7, col=LIGHT, align=PP_ALIGN.CENTER)

def top_bar(sl):
    rct(sl, 0, 0, W, 0.055, fill=ROSE_DARK)

def section_chip(sl, label):
    txt(sl, label, 0.5, 0.18, 9, 0.27, sz=7, bold=True, col=ROSE_DARK)

def slide_num(sl, n):
    txt(sl, f'{n} / 15', W-1.6, 0.18, 1.3, 0.27,
        sz=7, col=LIGHT, align=PP_ALIGN.RIGHT)

def content_title(sl, title, top=0.58):
    txt(sl, title, 0.5, top, 12.3, 0.6, sz=22, bold=True, col=DARK, font=FD)
    rct(sl, 0.5, top+0.63, 0.5, 0.045, fill=ROSE)

def embed_img(sl, buf, l, t, w, h):
    buf.seek(0)
    return sl.shapes.add_picture(buf, Inches(l), Inches(t), Inches(w), Inches(h))

# ── DIVIDER SLIDE FACTORY ─────────────────────────────────────────────────────

def divider_slide(num_str, section_word, title, desc):
    sl = new_slide(); set_bg(sl, BLUSH)
    # Faint huge background number
    txt(sl, num_str, -0.5, 0.1, 6, 6.8,
        sz=210, bold=True, col=FAINT_ROSE, font=FD)
    # Left accent bar
    rct(sl, 0, 0, 0.06, H, fill=ROSE_DARK)
    # Section word
    txt(sl, section_word, 1.2, 2.1, 9, 0.35, sz=8, bold=True, col=ROSE_DARK)
    # Accent dash
    rct(sl, 1.2, 2.53, 0.6, 0.05, fill=ROSE)
    # Title
    txt(sl, title, 1.2, 2.68, 10.8, 1.3, sz=36, bold=True, col=DARK, font=FD)
    # Description
    txt(sl, desc, 1.2, 4.2, 10.0, 0.9, sz=13, col=GRAY, font=FB)
    watermark(sl)
    return sl

# ── CHART GENERATORS ─────────────────────────────────────────────────────────

def chart_arousal():
    fig, ax = plt.subplots(figsize=(12, 4.2), dpi=150)
    fig.patch.set_facecolor('#fdfbf9'); ax.set_facecolor('#fdfbf9')

    x = np.linspace(0, 1, 600)
    y = (-0.03
         - 0.19 * np.exp(-((x - 0.415)**2) / 0.015)
         + 0.17 * np.exp(-((x - 0.875)**2) / 0.022)
         + 0.04 * x)

    bounds = [0, 0.28, 0.54, 0.75, 1.0]
    names  = ['Introduction', 'Methods', 'Results', 'Discussion']
    alphas = [0.06, 0.04, 0.06, 0.10]
    for i, (a, nm) in enumerate(zip(alphas, names)):
        ax.axvspan(bounds[i], bounds[i+1], alpha=a, color='#e8a4b4')

    ax.fill_between(x, y, y.min() - 0.004, alpha=0.18, color='#c4768a')
    ax.plot(x, y, color='#c4768a', linewidth=2.5, zorder=5)

    for b in bounds[1:-1]:
        ax.axvline(x=b, color='#e8a4b4', linewidth=1.2, linestyle='--', alpha=0.7)

    mi = np.argmin(y); ma = np.argmax(y)
    ax.plot(x[mi], y[mi], 'o', color='#c4768a', ms=8, zorder=6)
    ax.plot(x[ma], y[ma], 'o', color='#f5c8a0', ms=8,
            markeredgecolor='#c4768a', markeredgewidth=1.5, zorder=6)
    ax.annotate('Lowest arousal\n(Methods)',
                xy=(x[mi], y[mi]), xytext=(x[mi]-0.08, y[mi]-0.055),
                fontsize=9.5, color='#c4768a', ha='center',
                arrowprops=dict(arrowstyle='->', color='#c4768a', lw=1.2))
    ax.annotate('Peak arousal\n(Discussion)',
                xy=(x[ma], y[ma]), xytext=(x[ma]+0.07, y[ma]+0.045),
                fontsize=9.5, color='#8a6040', ha='center',
                arrowprops=dict(arrowstyle='->', color='#c4768a', lw=1.2))

    ym = y.min()
    for i, nm in enumerate(names):
        ax.text((bounds[i]+bounds[i+1])/2, ym-0.022, nm,
                ha='center', va='top', fontsize=10.5, color='#777')

    ax.set_xlim(0, 1)
    ax.set_ylabel('Arousal level (negative = calmer)', fontsize=10, color='#888', labelpad=6)
    ax.tick_params(labelbottom=False, bottom=False, labelsize=9, colors='#888')
    for s in ['top', 'right']:   ax.spines[s].set_visible(False)
    for s in ['bottom', 'left']: ax.spines[s].set_color('#e0e0e0')
    ax.set_title("Median Arousal Trajectory Across IMRaD Sections  |  Median Spearman's ρ = 0.758",
                 fontsize=11.5, color='#555', pad=12)
    plt.tight_layout(pad=0.5)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='#fdfbf9')
    plt.close(); buf.seek(0); return buf

def chart_kendall():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13.5, 3.6), dpi=150,
                                    gridspec_kw={'width_ratios': [1, 1], 'wspace': 0.35})
    fig.patch.set_facecolor('#fdfbf9')

    # Left: Kendall's W bars
    ax1.set_facecolor('#fdfbf9')
    dims = ['Valence', 'Dominance', 'Arousal']
    vals = [0.003, 0.159, 0.827]
    clrs = ['#cccccc', '#f5c8a0', '#c4768a']
    bars = ax1.barh(dims, vals, color=clrs, height=0.55, edgecolor='none')
    lbls = ["W = 0.003  (negligible)", "W = 0.159", "W = 0.827   ★★★"]
    for bar, lbl in zip(bars, lbls):
        ax1.text(bar.get_width() + 0.015, bar.get_y() + bar.get_height()/2,
                 lbl, va='center', fontsize=10, color='#444')
    ax1.set_xlim(0, 1.35); ax1.set_xlabel("Kendall's W (effect size)", fontsize=10, color='#888')
    ax1.tick_params(colors='#888', labelsize=10)
    for s in ['top', 'right']:   ax1.spines[s].set_visible(False)
    for s in ['bottom', 'left']: ax1.spines[s].set_color('#e0e0e0')
    ax1.grid(axis='x', alpha=0.3, color='#e0e0e0', zorder=0)
    ax1.set_title('Effect Size by VAD Dimension', fontsize=12, color='#555', pad=10)

    # Right: arousal section ranking
    ax2.set_facecolor('#fdfbf9')
    secs  = ['Methods', 'Results', 'Introduction', 'Discussion']
    arv   = [0.48, 0.65, 0.76, 0.95]
    clrs2 = ['#aaaaaa', '#f5c8a0', '#e8a4b4', '#c4768a']
    bars2 = ax2.barh(secs, arv, color=clrs2, height=0.55, edgecolor='none')
    for bar, lbl in zip(bars2, ['← Lowest', '', '', '← Highest']):
        if lbl:
            ax2.text(bar.get_width() + 0.01, bar.get_y() + bar.get_height()/2,
                     lbl, va='center', fontsize=10, color='#666', style='italic')
    ax2.set_xlim(0, 1.35); ax2.set_xlabel('Relative arousal level', fontsize=10, color='#888')
    ax2.tick_params(colors='#888', labelsize=10)
    for s in ['top', 'right']:   ax2.spines[s].set_visible(False)
    for s in ['bottom', 'left']: ax2.spines[s].set_color('#e0e0e0')
    ax2.grid(axis='x', alpha=0.3, color='#e0e0e0', zorder=0)
    ax2.set_title('Arousal Ranking Across IMRaD Sections', fontsize=12, color='#555', pad=10)
    ax2.text(0.5, -0.22, 'All 6 pairwise comparisons significant. Discussion > Introduction > Results > Methods.',
             transform=ax2.transAxes, ha='center', fontsize=8.5, color='#999', style='italic')

    plt.tight_layout(pad=0.8)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='#fdfbf9')
    plt.close(); buf.seek(0); return buf


# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 1 — TITLE
# ═════════════════════════════════════════════════════════════════════════════
sl = new_slide(); set_bg(sl, WHITE)
rct(sl, 0, 0, W, 0.055, fill=ROSE_DARK)          # top accent
rct(sl, 0, 0, 0.06, H,   fill=ROSE_DARK)          # left accent bar
ovl(sl, 9.0, -1.3, 5.8, 5.8, fill=BLUSH)          # decorative circle

txt(sl, 'CONFERENCE PRESENTATION  ·  2026',
    0.65, 0.85, 10, 0.35, sz=8, bold=True, col=ROSE_DARK)

txt(sl, 'Tracing Emotion Dynamics\nin Academic Writing',
    0.65, 1.3, 9.5, 2.0, sz=36, bold=True, col=DARK, font=FD)

txt(sl, 'A Probe into Medical Research Articles',
    0.65, 3.42, 9.5, 0.65, sz=20, italic=True, col=ROSE_DARK, font=FD)

# Divider
rct(sl, 0.65, 4.22, 0.55, 0.04, fill=BORDER)
ovl (sl, 1.24, 4.12, 0.08, 0.08, fill=ROSE)
rct(sl, 1.36, 4.22, 0.55, 0.04, fill=BORDER)

txt(sl, 'Ren Zhuoxuan', 0.65, 4.42, 7, 0.48, sz=16, bold=True, col=DARK, font=FD)
_, tf1 = txt(sl, 'Center for Research in Applied Linguistics', 0.65, 4.96, 10, 0.35, sz=11, col=GRAY)
add_para(tf1, 'Beijing Foreign Studies University', sz=11, col=GRAY, sa=0)
watermark(sl)

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 2 — TABLE OF CONTENTS
# ═════════════════════════════════════════════════════════════════════════════
sl = new_slide(); set_bg(sl, WARM)
rct(sl, 0, 0, W, 0.055, fill=ROSE)

txt(sl, 'CONTENTS', 0.8, 0.55, 5, 0.5, sz=9, bold=True, col=LIGHT, font=FB)

items = [('01', 'Research Questions'),
         ('02', 'Literature Review'),
         ('03', 'Research Design'),
         ('04', 'Results'),
         ('05', 'Implications')]

for i, (num, label) in enumerate(items):
    row_t = 1.25 + i * 1.1
    rct(sl, 0.8, row_t + 0.82, 11.7, 0.01, fill=BORDER)   # separator
    txt(sl, num,   0.8,  row_t, 1.0, 0.75, sz=26, bold=True, col=ROSE, font=FD)
    txt(sl, label, 2.0,  row_t + 0.14, 9.5, 0.55, sz=16, col=DARK, font=FB)

watermark(sl)

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 3 — DIVIDER: RESEARCH QUESTIONS
# ═════════════════════════════════════════════════════════════════════════════
divider_slide('01', 'SECTION ONE',
              'Research Questions',
              'What shape does emotion take as it unfolds across a medical research article?')

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 4 — RESEARCH QUESTIONS CONTENT
# ═════════════════════════════════════════════════════════════════════════════
sl = new_slide(); set_bg(sl, WHITE)
top_bar(sl); section_chip(sl, '01 — RESEARCH QUESTIONS'); slide_num(sl, 4)
content_title(sl, 'Motivation & Questions')

# Gap / motivation box
rct(sl, 0.5, 1.35, 12.3, 1.0, fill=BLUSH, line_rgb=BORDER)
rct(sl, 0.5, 1.35, 0.06, 1.0, fill=ROSE)   # left accent
txt(sl, 'Emotion arcs have been studied in novels and films — but academic discourse has received little '
        'attention from this perspective. Most existing studies represent emotion as simply positive or negative. '
        'The VAD dimensional model offers a finer-grained analysis, yet has rarely been applied to academic writing.',
    0.72, 1.42, 11.9, 0.82, sz=11, col=GRAY, font=FB)

# RQ 1
rct(sl, 0.5, 2.55, 12.3, 1.0, fill=CARD, line_rgb=BORDER)
ovl(sl, 0.55, 2.65, 0.65, 0.65, fill=ROSE_DARK)
txt(sl, 'RQ 1', 0.55, 2.65, 0.65, 0.65, sz=8, bold=True, col=WHITE,
    align=PP_ALIGN.CENTER, font=FB)
txt(sl, 'Do medical research articles exhibit identifiable emotion trajectories '
        'across the three VAD dimensions (Valence, Arousal, Dominance)?',
    1.35, 2.68, 11.3, 0.75, sz=12, col=DARK, font=FB)

# RQ 2
rct(sl, 0.5, 3.7, 12.3, 1.0, fill=CARD, line_rgb=BORDER)
ovl(sl, 0.55, 3.8, 0.65, 0.65, fill=ROSE_DARK)
txt(sl, 'RQ 2', 0.55, 3.8, 0.65, 0.65, sz=8, bold=True, col=WHITE,
    align=PP_ALIGN.CENTER, font=FB)
txt(sl, 'Do these trajectories differ systematically across the four IMRaD sections — '
        'Introduction, Methods, Results, and Discussion?',
    1.35, 3.83, 11.3, 0.75, sz=12, col=DARK, font=FB)

watermark(sl)

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 5 — DIVIDER: LITERATURE REVIEW
# ═════════════════════════════════════════════════════════════════════════════
divider_slide('02', 'SECTION TWO',
              'Literature Review',
              'What do we know about emotion in academic writing — and where are the gaps?')

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 6 — LITERATURE REVIEW CONTENT
# ═════════════════════════════════════════════════════════════════════════════
sl = new_slide(); set_bg(sl, WHITE)
top_bar(sl); section_chip(sl, '02 — LITERATURE REVIEW'); slide_num(sl, 6)
content_title(sl, 'What We Know — and the Gaps')

cards = [
    ('1.  Positivity Bias',
     'Studies document a linguistic positivity bias across academic genres — '
     'a tendency toward positive expression that has strengthened over time.'),
    ('2.  Section-Level Variation',
     'Sentiment scores vary across discourse units (sections, moves), suggesting '
     'emotion is patterned by communicative function.'),
    ('3.  Methodological Limit',
     'Existing studies assign a single score to a predefined unit — they cannot '
     'capture how emotion shifts continuously over an article.'),
    ('4.  The VAD Model',
     'Valence (pleasant↔unpleasant) · Arousal (activated↔calm) · Dominance '
     '(in-control↔submissive) — rarely applied to academic writing.'),
]
cols = [0.5, 6.65]
for i, (heading, body) in enumerate(cards):
    col = cols[i % 2]; row = 1.35 if i < 2 else 3.1
    rct(sl, col, row, 5.95, 1.55, fill=CARD, line_rgb=BORDER)
    txt(sl, heading, col+0.15, row+0.1, 5.6, 0.35, sz=10, bold=True, col=ROSE_DARK)
    txt(sl, body,    col+0.15, row+0.45, 5.6, 1.0,  sz=10, col=GRAY)

# Gap box
rct(sl, 0.5, 4.82, 12.3, 0.65, fill=BLUSH, line_rgb=BORDER)
rct(sl, 0.5, 4.82, 0.06, 0.65, fill=ROSE_DARK)
txt(sl, 'Research Gap:  Emotion arcs have been identified in novels and films — but academic discourse '
        'has received little attention, and the VAD model remains underused in genre analysis.',
    0.72, 4.9, 11.9, 0.5, sz=10.5, bold=True, col=ROSE_DARK)

watermark(sl)

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 7 — DIVIDER: RESEARCH DESIGN
# ═════════════════════════════════════════════════════════════════════════════
divider_slide('03', 'SECTION THREE',
              'Research Design',
              'Corpus, lexicon, rolling-window method, and statistical framework.')

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 8 — RESEARCH DESIGN CONTENT
# ═════════════════════════════════════════════════════════════════════════════
sl = new_slide(); set_bg(sl, WHITE)
top_bar(sl); section_chip(sl, '03 — RESEARCH DESIGN'); slide_num(sl, 8)
content_title(sl, 'Corpus & Method')

# Left column header
txt(sl, 'CORPUS', 0.5, 1.42, 5.9, 0.28, sz=7.5, bold=True, col=ROSE_DARK)
rct(sl, 0.5, 1.72, 5.9, 0.015, fill=BORDER)

# Corpus stats
for val, label, ty in [
    ('3,600', 'medical research articles\nIMRaD structure, auto-extracted', 1.82),
    ('13.6 M', 'total tokens across the corpus', 2.72),
]:
    txt(sl, val,   0.5,  ty,       2.0, 0.55, sz=24, bold=True, col=ROSE_DARK, font=FD)
    txt(sl, label, 2.65, ty+0.1,   3.8, 0.55, sz=10, col=GRAY)

# Statistical tests sub-header
txt(sl, 'STATISTICAL TESTS', 0.5, 3.55, 5.9, 0.28, sz=7.5, bold=True, col=ROSE_DARK)
rct(sl, 0.5, 3.85, 5.9, 0.015, fill=BORDER)

stat_steps = [
    'RQ1:  Spearman\'s rho — each article vs. corpus median trajectory',
    'RQ2:  Friedman test + Kendall\'s W (effect size) — section-level VAD differences',
]
for j, step in enumerate(stat_steps):
    ovl(sl, 0.5, 3.98+j*0.65, 0.3, 0.3, fill=ROSE)
    txt(sl, 'AB'[j], 0.5, 3.98+j*0.65, 0.3, 0.3, sz=8, bold=True, col=WHITE,
        align=PP_ALIGN.CENTER)
    txt(sl, step, 0.95, 4.0+j*0.65, 5.45, 0.45, sz=10.5, col=GRAY)

# Right column header
txt(sl, 'ROLLING-WINDOW PROCEDURE', 6.8, 1.42, 6.0, 0.28, sz=7.5, bold=True, col=ROSE_DARK)
rct(sl, 6.8, 1.72, 6.0, 0.015, fill=BORDER)

method_steps = [
    'Score each word using the NRC-VAD Lexicon on V, A, D scales',
    'Slide a 500-token window one token at a time; compute average VAD at each position',
    'Normalize each article\'s trajectory to a common scale (0–100%)',
    'Map onto the IMRaD structure → three continuous emotion curves per article',
]
for j, step in enumerate(method_steps):
    ovl(sl, 6.8, 1.85+j*1.15, 0.38, 0.38, fill=ROSE_DARK)
    txt(sl, str(j+1), 6.8, 1.85+j*1.15, 0.38, 0.38, sz=10, bold=True, col=WHITE,
        align=PP_ALIGN.CENTER)
    txt(sl, step, 7.32, 1.87+j*1.15, 5.45, 0.7, sz=11, col=GRAY)

# Divider between columns
rct(sl, 6.65, 1.35, 0.02, 5.6, fill=BORDER)
watermark(sl)

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 9 — DIVIDER: RESULTS
# ═════════════════════════════════════════════════════════════════════════════
divider_slide('04', 'SECTION FOUR',
              'Results',
              'Valence and Dominance stay positive — but Arousal tells a different, highly consistent story.')

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 10 — RESULTS OVERVIEW
# ═════════════════════════════════════════════════════════════════════════════
sl = new_slide(); set_bg(sl, WHITE)
top_bar(sl); section_chip(sl, '04 — RESULTS'); slide_num(sl, 10)
content_title(sl, 'VAD Trajectories: Overview')

for col_l, panel_title, badge, bullets, highlight in [
    (0.5,
     'Valence & Dominance', 'V & D',
     ['Both remain positive throughout — consistent with the documented positivity bias',
      'Cross-article shape consistency is low',
      'Median Spearman\'s ρ for Valence = 0.064'],
     'Individual articles do NOT share a common trajectory shape on V and D. No identifiable arc.'),
    (7.0,
     'Arousal', 'A ★',
     ['Remains negative throughout — calm, low-activation register',
      'Traces a clear fall-rise arc: lowest in Methods, peak in Discussion',
      'Median Spearman\'s ρ = 0.758 — strong cross-article consistency'],
     'This arc is reliably reproduced across individual articles, aligning closely with IMRaD boundaries.'),
]:
    rct(sl, col_l, 1.35, 5.8, 5.25, fill=CARD, line_rgb=BORDER)
    txt(sl, panel_title, col_l+0.18, 1.48, 4.0, 0.4, sz=13, bold=True, col=DARK, font=FD)
    rct(sl, col_l+4.55, 1.55, 0.9, 0.28, fill=BLUSH)
    txt(sl, badge, col_l+4.55, 1.55, 0.9, 0.28, sz=7.5, bold=True, col=ROSE_DARK,
        align=PP_ALIGN.CENTER)
    rct(sl, col_l+0.18, 1.9, 5.35, 0.015, fill=BORDER)
    for k, b in enumerate(bullets):
        txt(sl, '›  ' + b, col_l+0.18, 2.05+k*0.62, 5.4, 0.55, sz=10.5, col=GRAY)
    rct(sl, col_l+0.18, 3.95, 5.4, 0.65, fill=BLUSH, line_rgb=BORDER)
    txt(sl, highlight, col_l+0.3, 4.02, 5.1, 0.5, sz=10, col=ROSE_DARK, bold=True)

watermark(sl)

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 11 — AROUSAL ARC CHART
# ═════════════════════════════════════════════════════════════════════════════
sl = new_slide(); set_bg(sl, WHITE)
top_bar(sl); section_chip(sl, '04 — RESULTS'); slide_num(sl, 11)
content_title(sl, 'The Arousal Arc Across IMRaD Sections')

arc_buf = chart_arousal()
embed_img(sl, arc_buf, 0.4, 1.35, 12.5, 5.5)
watermark(sl)

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 12 — SECTION DIFFERENCES (Kendall's W)
# ═════════════════════════════════════════════════════════════════════════════
sl = new_slide(); set_bg(sl, WHITE)
top_bar(sl); section_chip(sl, '04 — RESULTS'); slide_num(sl, 12)
content_title(sl, 'Section-Level Differences (Friedman Tests)')

kw_buf = chart_kendall()
embed_img(sl, kw_buf, 0.35, 1.35, 12.6, 5.5)
watermark(sl)

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 13 — DIVIDER: IMPLICATIONS
# ═════════════════════════════════════════════════════════════════════════════
divider_slide('05', 'SECTION FIVE',
              'Implications',
              'What these findings mean for theory, methodology, and language teaching.')

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 14 — IMPLICATIONS CONTENT
# ═════════════════════════════════════════════════════════════════════════════
sl = new_slide(); set_bg(sl, WHITE)
top_bar(sl); section_chip(sl, '05 — IMPLICATIONS'); slide_num(sl, 14)
content_title(sl, 'Three Layers of Significance')

impl_cols = [
    ('🔬', 'THEORETICAL',
     ['VAD dimensions capture distinct aspects of affect — they are not interchangeable',
      'Arousal is sensitive to genre-internal functional variation',
      'Dimensional models offer a richer lens than polarity analysis']),
    ('⚙', 'METHODOLOGICAL',
     ['Rolling window + time normalization + IMRaD mapping = replicable framework',
      'Applicable across disciplines, genres, historical periods',
      'Adapts the emotion-arc approach from fiction to academic discourse']),
    ('📚', 'PEDAGOGICAL',
     ['Different IMRaD sections carry different affective expectations',
      'Guide students to manage emotional resources deliberately across sections',
      'High-contribution words → basis for an ESP-oriented pedagogical lexicon']),
]
for i, (icon, heading, bullets) in enumerate(impl_cols):
    cl = 0.5 + i * 4.3
    rct(sl, cl, 1.35, 4.0, 5.3, fill=CARD, line_rgb=BORDER)
    # Icon box
    rct(sl, cl+0.15, 1.5,  0.55, 0.55, fill=BLUSH)
    txt(sl, icon, cl+0.15, 1.5, 0.55, 0.55, sz=16, align=PP_ALIGN.CENTER)
    txt(sl, heading, cl+0.15, 2.18, 3.7, 0.35, sz=8, bold=True, col=ROSE_DARK)
    rct(sl, cl+0.15, 2.57, 3.7, 0.015, fill=BORDER)
    for k, b in enumerate(bullets):
        txt(sl, '›  ' + b, cl+0.15, 2.72+k*0.82, 3.7, 0.72, sz=10.5, col=GRAY)

watermark(sl)

# ═════════════════════════════════════════════════════════════════════════════
#  SLIDE 15 — THANK YOU
# ═════════════════════════════════════════════════════════════════════════════
sl = new_slide(); set_bg(sl, BLUSH)

penguin_path = '/home/user/frontend-slides/.claude-design/pptx-extract/assets/slide14_img1.png'
with open(penguin_path, 'rb') as f:
    pbuf = io.BytesIO(f.read())
embed_img(sl, pbuf, 5.0, 0.3, 3.35, 4.3)

txt(sl, 'Thank You', 0.0, 4.7, W, 0.9,
    sz=40, bold=True, col=DARK, font=FD, align=PP_ALIGN.CENTER)
txt(sl, 'Any questions, comments, or suggestions are most welcome.',
    0.0, 5.72, W, 0.48,
    sz=14, col=GRAY, align=PP_ALIGN.CENTER)
txt(sl, 'Ren Zhuoxuan  ·  corpus.bfsu.edu.cn',
    0.0, 6.35, W, 0.38,
    sz=11, col=LIGHT, align=PP_ALIGN.CENTER)

# ═════════════════════════════════════════════════════════════════════════════
#  SAVE
# ═════════════════════════════════════════════════════════════════════════════
out = '/home/user/frontend-slides/emotion-dynamics-slides.pptx'
prs.save(out)
print(f'Saved: {out}  ({len(prs.slides)} slides)')
