# -*- coding: utf-8 -*-
"""
STELOS - The ASD-STE100 Intelligence Platform (an ALTEN product).
Simplified Technical English (ASD-STE100 Issue 9) checker
with in-place word replacement into a COPY of the source file.
RR / ALTEN internal tool. Bind 127.0.0.1, port 5002.

Inputs supported: pasted text, .docx, .txt/.md, .xlsx/.xlsm
Replacement:  tick the words to fix -> tool writes a corrected COPY and offers it
              for download. The original uploaded file is NEVER modified.

Ships with ste_dictionary.json (ASD-STE100 Issue 9, 2025-01-15) in the SAME folder.
"""
import os, re, json, threading, webbrowser, time, traceback, html, uuid, tempfile
from datetime import datetime
from flask import Flask, request, jsonify, render_template_string, send_file

app = Flask(__name__)
PORT = 5002
HERE = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(HERE, 'ste_dictionary.json')
TMPDIR = os.path.join(tempfile.gettempdir(), 'ste_checker')
os.makedirs(TMPDIR, exist_ok=True)

# ---------------------------------------------------------------- load data
with open(DATA_PATH, 'r', encoding='utf-8') as f:
    STE = json.load(f)
APPROVED = STE['approved']
NOT_APPROVED = STE['not_approved']
RULES = STE['rules']
META = STE['meta']

ALWAYS_OK = set("a an the and or of to in on at for from with by as is are was were "
                "be do this that it not no if".split())

# token registry: token -> {path, filename, ext}
UPLOADS = {}

# ---- shared technical-terms library ("DB" file, safe for a network share) ----
# Point STE_TERMS_FILE at a shared drive so all users read/write the same list.
TERMS_FILE = os.environ.get('STE_TERMS_FILE', os.path.join(HERE, 'tech_terms.json'))
TERMS_LOCK = threading.Lock()
CURRENT_USER = os.environ.get('USERNAME') or os.environ.get('USER') or 'user'

def load_terms():
    """Read the shared library fresh from disk so other users' additions show."""
    try:
        with open(TERMS_FILE, 'r', encoding='utf-8') as fh:
            d = json.load(fh)
            return d if isinstance(d, dict) else {}
    except (FileNotFoundError, ValueError, OSError):
        return {}

def add_terms(words, by):
    """Merge new approved technical terms into the shared file (atomic write)."""
    with TERMS_LOCK:
        d = load_terms()
        added = []
        for w in words:
            wl = (w or '').strip().lower()
            if wl and wl not in d:
                d[wl] = {'by': by, 'at': datetime.now().strftime('%Y-%m-%d %H:%M')}
                added.append(wl)
        if added:
            tmp = TERMS_FILE + '.tmp'
            with open(tmp, 'w', encoding='utf-8') as fh:
                json.dump(d, fh, indent=2, ensure_ascii=False)
            os.replace(tmp, TERMS_FILE)
        return d, added

# ---------------------------------------------------------------- word logic
def deinflect(w):
    forms = {w}
    for suf, repl in (('ies','y'), ('ied','y'), ('ing',''), ('ed',''),
                      ('es',''), ('s',''), ('ly','')):
        if w.endswith(suf) and len(w) - len(suf) >= 2:
            base = w[:-len(suf)] + repl
            forms.add(base); forms.add(base + 'e')
    return forms

def lookup_not_approved(word):
    for f in deinflect(word):
        if f in NOT_APPROVED:
            return NOT_APPROVED[f]
    return None

def is_approved(word):
    return any(f in APPROVED for f in deinflect(word))

# ---------------------------------------------------------------- rule engine
BRITISH = {
 'colour':'color','colours':'colors','behaviour':'behavior','behaviours':'behaviors',
 'centre':'center','centres':'centers','metre':'meter','metres':'meters','litre':'liter',
 'litres':'liters','fibre':'fiber','fibres':'fibers','defence':'defense','licence':'license',
 'catalogue':'catalog','catalogues':'catalogs','programme':'program','programmes':'programs',
 'analyse':'analyze','analysed':'analyzed','organise':'organize','organised':'organized',
 'realise':'realize','minimise':'minimize','maximise':'maximize','initialise':'initialize',
 'aluminium':'aluminum','tyre':'tire','tyres':'tires','grey':'gray','mould':'mold',
 'modelling':'modeling','labelling':'labeling','travelling':'traveling','cancelled':'canceled',
 'fuelled':'fueled','signalling':'signaling','practise':'practice','sulphur':'sulfur',
 'manoeuvre':'maneuver','oxidised':'oxidized','optimise':'optimize','favour':'favor',
}
CONTRACTIONS = {
 "don't":'do not',"doesn't":'does not',"didn't":'did not',"isn't":'is not',"aren't":'are not',
 "wasn't":'was not',"weren't":'were not',"can't":'cannot',"couldn't":'could not',
 "won't":'will not',"wouldn't":'would not',"shouldn't":'should not',"mustn't":'must not',
 "haven't":'have not',"hasn't":'has not',"hadn't":'had not',"it's":'it is',"that's":'that is',
 "there's":'there is',"we're":'we are',"you're":'you are',"they're":'they are',"i'm":'i am',
 "we've":'we have',"you've":'you have',"they've":'they have',"i've":'i have',"we'll":'we will',
 "you'll":'you will',"they'll":'they will',"it'll":'it will',"let's":'let us',"who's":'who is',
}
PHRASAL = {
 'set up':'install / prepare','carry out':'do','put out':'extinguish / remove','shut down':'stop',
 'pull out':'remove','take off':'remove','look for':'find','find out':'find','go through':'examine',
}
BE = {'is','are','was','were','be','been','being','am'}
HAVE = {'have','has','had'}
IRREG_PP = {'done','made','given','taken','seen','shown','set','put','cut','built','held','kept',
 'sent','found','removed','installed','completed','opened','closed','connected','disconnected',
 'gone','known','written','driven','broken','worn','torn','begun','shut','read','told'}

def _is_pp(w):
    return w.endswith('ed') or w in IRREG_PP

def analyze_constructs(s, mode):
    """Sentence-level rule findings that are NOT simple word swaps.
    Returns list of (rule, severity, message)."""
    out = []
    toks = WORD_RE.findall(s)
    low = [t.lower() for t in toks]
    if ';' in s:
        out.append(('8.1', 'error', 'Semicolon (;) is not allowed in STE.'))
    sl = ' ' + ' '.join(low) + ' '
    for ph, alt in PHRASAL.items():
        if ' ' + ph + ' ' in sl:
            out.append(('9.3', 'advisory', 'Possible phrasal verb "%s" - consider "%s".' % (ph, alt)))
    for i in range(len(low) - 1):
        a, b = low[i], low[i + 1]
        if a in BE and b.endswith('ing') and len(b) > 4:
            out.append(('3.5', 'warn', 'Progressive form "%s %s" - use the simple/imperative form.' % (a, b)))
        elif a in BE and _is_pp(b) and mode == 'procedural':
            out.append(('3.6', 'warn', 'Passive voice "%s %s" - use the active voice in procedures.' % (a, b)))
        if a in HAVE and _is_pp(b):
            out.append(('3.4', 'warn', 'Perfect tense "%s %s" - use the simple past tense.' % (a, b)))
    if mode == 'procedural' and low:
        if low[0] in ('you', 'we', 'the', 'this', 'these', 'it') and len(low) > 1 and \
           low[1] in ('must', 'should', 'will', 'can', 'shall', 'need', 'have', 'has'):
            out.append(('5.3', 'advisory', 'Write the instruction in the command (imperative) form.'))
    return out

# Full ASD-STE100 Issue 9 rule catalogue. mode: 'auto' = tool checks it,
# 'manual' = needs human judgement (listed so coverage is complete).
RULE_CATALOG = [
 ('Section 1 - Words', [
   ('1.1','Use approved words','auto'),
   ('1.2','Use words only as the specified part of speech','manual'),
   ('1.3','Use approved words only with their approved meanings','manual'),
   ('1.4','Use only approved forms of verbs and adjectives','auto'),
   ('1.5','Use words that fit a technical noun category','manual'),
   ('1.6','Use a non-approved word only as a technical noun','auto'),
   ('1.7','Do not use technical nouns as verbs','manual'),
   ('1.8','Use technical nouns approved in your company/industry','manual'),
   ('1.9','Select technical nouns that are short and clear','manual'),
   ('1.10','No regional, slang, or jargon technical nouns','manual'),
   ('1.11','Do not use different nouns for the same item','manual'),
   ('1.12','Use verbs that fit a technical verb category','manual'),
   ('1.13','Do not use technical verbs as nouns','manual'),
   ('1.14','Use American English spelling','auto'),
 ]),
 ('Section 2 - Multi-word nouns', [
   ('2.1','Multi-word nouns of no more than three words','manual'),
   ('2.2','Write nouns of more than three words in full','manual'),
 ]),
 ('Section 3 - Verbs', [
   ('3.1','Use only verb forms given in the dictionary','auto'),
   ('3.2','Use only the allowed verb forms and tenses','auto'),
   ('3.3','Use the past participle as an adjective','manual'),
   ('3.4','No auxiliary verbs for complex constructions','auto'),
   ('3.5','Use "-ing" only as a technical noun or modifier','auto'),
   ('3.6','Use the active voice (passive only in descriptive)','auto'),
   ('3.7','Use a verb, not a noun, to describe an action','manual'),
 ]),
 ('Section 4 - Sentences', [
   ('4.1','Write short and clear sentences','auto'),
   ('4.2','No omitted words or contractions','auto'),
   ('4.3','Use a vertical list for complex text','manual'),
   ('4.4','Use connecting words and phrases','manual'),
   ('4.5','Use an article or demonstrative adjective','manual'),
 ]),
 ('Section 5 - Procedural writing', [
   ('5.1','Maximum 20 words in a procedural sentence','auto'),
   ('5.2','Only one instruction in each sentence','manual'),
   ('5.3','Write instructions in the imperative form','auto'),
   ('5.4','Start with the condition, then the instruction','manual'),
   ('5.5','Write notes to give information, not instructions','manual'),
 ]),
 ('Section 6 - Descriptive writing', [
   ('6.1','Give information gradually','manual'),
   ('6.2','Use key words and phrases for structure','manual'),
   ('6.3','Maximum 25 words in a descriptive sentence','auto'),
   ('6.4','Use paragraphs to show related information','manual'),
   ('6.5','Each paragraph has only one topic','manual'),
   ('6.6','No paragraph has more than six sentences','auto'),
 ]),
 ('Section 7 - Safety instructions', [
   ('7.1','Identify with "warning" or "caution"','manual'),
   ('7.2','Start a safety instruction with a command','manual'),
   ('7.3','Give an explanation of the risk','manual'),
 ]),
 ('Section 8 - Punctuation and word count', [
   ('8.1','No semicolon; standard punctuation only','auto'),
   ('8.2','Use hyphens to connect related words','manual'),
   ('8.3','Use parentheses only as permitted','manual'),
   ('8.4','Colon in a vertical list counts like a period','manual'),
   ('8.5','Text in parentheses counts as one word','manual'),
   ('8.6','Count listed elements as one word','manual'),
   ('8.7','Hyphenated words count as one word','manual'),
 ]),
 ('Section 9 - Word usage', [
   ('9.1','Use a different construction when a swap fails','manual'),
   ('9.2','Use each approved word correctly','manual'),
   ('9.3','Do not make phrasal verbs','auto'),
   ('9.4','Use a consistent style and terminology','manual'),
 ]),
]

# ---------------------------------------------------------------- text parse
SENT_SPLIT = re.compile(r'(?<=[.!?:])\s+')
WORD_RE = re.compile(r"[A-Za-z][A-Za-z'\-]*")

def split_sentences(text):
    out = []
    for para_i, para in enumerate(re.split(r'\n\s*\n', text)):
        para = para.strip()
        if not para:
            continue
        sents = [s.strip() for s in SENT_SPLIT.split(para) if s.strip()]
        out.append((para_i, sents))
    return out

def default_replacement(alt):
    """Turn dictionary alternative into a natural body-text default."""
    if not alt or alt.startswith('Not in STE'):
        return ''
    return alt.lower()

def in_allowed(word, extra_allowed):
    return any(f in extra_allowed for f in deinflect(word))

def check_text(text, mode, strict, extra_allowed=None):
    extra_allowed = extra_allowed or set()
    sent_limit = RULES['proc_sentence_max'] if mode == 'procedural' else RULES['desc_sentence_max']
    issues, flagged_words = [], {}
    rule_findings = []          # {rule, severity, msg, context}
    rule_counts = {}            # rule number -> count
    total_words = total_sentences = 0
    paras = split_sentences(text)

    def add_finding(rule, sev, msg, context):
        rule_findings.append({'rule': rule, 'severity': sev, 'msg': msg, 'context': context})
        rule_counts[rule] = rule_counts.get(rule, 0) + 1

    for para_i, sents in paras:
        if len(sents) > RULES['para_max_sentences']:
            issues.append({'type': 'paragraph', 'severity': 'warn',
                           'msg': 'Paragraph has %d sentences (max %d).' %
                                  (len(sents), RULES['para_max_sentences']),
                           'context': (sents[0][:80] + '...') if sents else ''})
            add_finding('6.6', 'warn', 'Paragraph has %d sentences (max %d).' %
                        (len(sents), RULES['para_max_sentences']),
                        (sents[0][:80] + '...') if sents else '')
        for s in sents:
            total_sentences += 1
            ctx = s[:110] + ('...' if len(s) > 110 else '')
            words = WORD_RE.findall(s)
            total_words += len(words)
            if len(words) > sent_limit:
                issues.append({'type': 'length', 'severity': 'warn',
                               'msg': 'Sentence has %d words (max %d for %s writing).' %
                                      (len(words), sent_limit, mode), 'context': ctx})
                add_finding('5.1' if mode == 'procedural' else '6.3', 'warn',
                            'Sentence has %d words (max %d).' % (len(words), sent_limit), ctx)
            # sentence-level construct rules (verbs, semicolon, phrasal, imperative)
            for rule, sev, msg in analyze_constructs(s, mode):
                add_finding(rule, sev, msg, ctx)
            # word-level rules that ARE fixable word swaps
            for w in words:
                wl = w.lower()
                if wl in CONTRACTIONS:
                    e = flagged_words.setdefault('c:' + wl,
                        {'word': w, 'alt': CONTRACTIONS[wl], 'pos': '4.2', 'count': 0,
                         'has_alt': True, 'replacement': CONTRACTIONS[wl]})
                    e['count'] += 1
                    rule_counts['4.2'] = rule_counts.get('4.2', 0) + 1
                    continue
                if wl in BRITISH and BRITISH[wl] != wl:
                    e = flagged_words.setdefault('b:' + wl,
                        {'word': w, 'alt': BRITISH[wl], 'pos': '1.14', 'count': 0,
                         'has_alt': True, 'replacement': BRITISH[wl]})
                    e['count'] += 1
                    rule_counts['1.14'] = rule_counts.get('1.14', 0) + 1
                    continue
                if wl in ALWAYS_OK:
                    continue
                na = lookup_not_approved(wl)
                if na:
                    e = flagged_words.setdefault(na['word'].lower(),
                        {'word': na['word'], 'alt': na['alt'] or '(rewrite required)',
                         'pos': na['pos'], 'count': 0, 'has_alt': bool(na['alt']),
                         'replacement': default_replacement(na['alt'])})
                    e['count'] += 1
                    rule_counts['1.1'] = rule_counts.get('1.1', 0) + 1
                elif (strict and not is_approved(wl) and not w[0].isupper()
                      and not in_allowed(wl, extra_allowed)):
                    e = flagged_words.setdefault('~' + wl,
                        {'word': w, 'alt': 'Not in STE dictionary - verify it is an '
                         'approved technical noun/verb or reword.', 'pos': '?', 'count': 0,
                         'has_alt': False, 'replacement': ''})
                    e['count'] += 1
                    rule_counts['1.6'] = rule_counts.get('1.6', 0) + 1

    flagged = sorted(flagged_words.values(), key=lambda x: (-x['count'], x['word'].lower()))
    rule_findings.sort(key=lambda x: (x['rule']))

    # coverage report over the whole standard
    coverage = []
    auto_checked = auto_total = 0
    for section, rules in RULE_CATALOG:
        srules = []
        for num, title, kind in rules:
            cnt = rule_counts.get(num, 0)
            if kind == 'auto':
                auto_total += 1
                if cnt:
                    auto_checked += 1
            srules.append({'num': num, 'title': title, 'kind': kind, 'count': cnt})
        coverage.append({'section': section, 'rules': srules})

    summary = {
        'words': total_words, 'sentences': total_sentences, 'paragraphs': len(paras),
        'unapproved_total': sum(f['count'] for f in flagged),
        'length_issues': len([i for i in issues if i['type'] == 'length']),
        'paragraph_issues': len([i for i in issues if i['type'] == 'paragraph']),
        'rule_issues': len(rule_findings),
        'auto_rules': auto_total, 'mode': mode, 'sentence_limit': sent_limit,
    }
    return {'summary': summary, 'flagged': flagged, 'issues': issues,
            'rule_findings': rule_findings, 'coverage': coverage,
            'highlighted': build_highlight(text, flagged_words, strict, extra_allowed)}

def build_highlight(text, flagged_words, strict, extra_allowed=None):
    extra_allowed = extra_allowed or set()
    flagged_bases = {e['word'].lower() for e in flagged_words.values()}
    def repl(m):
        w = m.group(0); wl = w.lower()
        if wl in CONTRACTIONS:
            return '<mark class="bad" title="Write: %s">%s</mark>' % (
                html.escape(CONTRACTIONS[wl]), html.escape(w))
        if wl in BRITISH and BRITISH[wl] != wl:
            return '<mark class="bad" title="American: %s">%s</mark>' % (
                html.escape(BRITISH[wl]), html.escape(w))
        if in_allowed(wl, extra_allowed):
            return '<mark class="good" title="Approved technical term">%s</mark>' % html.escape(w)
        hit = lookup_not_approved(wl)
        if hit and hit['word'].lower() in flagged_bases:
            return '<mark class="bad" title="Use: %s">%s</mark>' % (
                html.escape(hit['alt'] or 'rewrite'), html.escape(w))
        if (strict and not is_approved(wl) and wl not in ALWAYS_OK
                and not w[0].isupper() and not in_allowed(wl, extra_allowed)):
            return '<mark class="warnw" title="Not in STE dictionary">%s</mark>' % html.escape(w)
        return html.escape(w)
    out, last = [], 0
    for m in WORD_RE.finditer(text):
        out.append(html.escape(text[last:m.start()])); out.append(repl(m)); last = m.end()
    out.append(html.escape(text[last:]))
    return ''.join(out).replace('\n', '<br>')

# ---------------------------------------------------------------- readers
def extract_text_from(path, ext):
    if ext == '.docx':
        from docx import Document
        doc = Document(path)
        parts = [p.text for p in doc.paragraphs]
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    parts.append(cell.text)
        return '\n\n'.join(parts)
    if ext in ('.txt', '.md'):
        with open(path, 'rb') as fh:
            return fh.read().decode('utf-8', errors='replace')
    if ext in ('.xlsx', '.xlsm'):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        parts = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    if isinstance(v, str) and v.strip():
                        parts.append(v)
        wb.close()
        return '\n\n'.join(parts)
    raise ValueError('Unsupported file type: ' + ext)

# ---------------------------------------------------------------- replacement
def cased(orig, repl):
    """Match the capitalization of the original occurrence."""
    if orig.isupper():
        return repl.upper()
    if orig[:1].isupper():
        return repl[:1].upper() + repl[1:]
    return repl

def build_patterns(pairs):
    """pairs: list of (find_word, replace_with). Returns compiled whole-word patterns."""
    pats = []
    for find, rep in pairs:
        if not find or not rep:
            continue
        pats.append((re.compile(r"(?<![A-Za-z'])" + re.escape(find) + r"(?![A-Za-z'])",
                                 re.IGNORECASE), rep))
    return pats

def sub_text(text, pats, counter):
    for pat, rep in pats:
        def _r(m):
            counter[0] += 1
            return cased(m.group(0), rep)
        text = pat.sub(_r, text)
    return text

def replace_in_txt(src, dst, pats):
    counter = [0]
    with open(src, 'rb') as fh:
        text = fh.read().decode('utf-8', errors='replace')
    text = sub_text(text, pats, counter)
    with open(dst, 'w', encoding='utf-8') as fh:
        fh.write(text)
    return counter[0]

def _replace_paragraph(p, pats, counter):
    """Apply replacements to a paragraph exactly once.
    Uses run-level replacement (keeps inline formatting) when every match sits
    inside a single run; falls back to a merged replacement only when a word is
    split across runs. This avoids double-applying when the replacement text
    itself contains the original word (e.g. motor -> dry-motor)."""
    runs = p.runs
    if not runs:
        return
    texts = [r.text for r in runs]
    joined = ''.join(texts)
    if not joined:
        return
    bounds, pos = [], 0
    for t in texts:
        bounds.append((pos, pos + len(t))); pos += len(t)
    cross = False
    for pat, _rep in pats:
        for mm in pat.finditer(joined):
            s, e = mm.span()
            if not any(bs <= s and e <= be for bs, be in bounds):
                cross = True; break
        if cross:
            break
    if cross:
        runs[0].text = sub_text(joined, pats, counter)
        for r in runs[1:]:
            r.text = ''
    else:
        for r in runs:
            if r.text:
                r.text = sub_text(r.text, pats, counter)

def replace_in_docx(src, dst, pats):
    from docx import Document
    counter = [0]
    doc = Document(src)
    for p in doc.paragraphs:
        _replace_paragraph(p, pats, counter)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_paragraph(p, pats, counter)
    # headers / footers too
    for section in doc.sections:
        for hf in (section.header, section.footer):
            for p in hf.paragraphs:
                _replace_paragraph(p, pats, counter)
    doc.save(dst)
    return counter[0]

def replace_in_xlsx(src, dst, pats):
    import openpyxl
    counter = [0]
    wb = openpyxl.load_workbook(src)  # keep_vba false; formatting preserved
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value:
                    new = sub_text(cell.value, pats, counter)
                    if new != cell.value:
                        cell.value = new
    wb.save(dst)
    return counter[0]

# ---------------------------------------------------------------- routes
@app.route('/')
def index():
    return render_template_string(PAGE, meta=META, rules=RULES)

@app.route('/check', methods=['POST'])
def check():
    try:
        mode = request.form.get('mode', 'procedural')
        strict = request.form.get('strict', 'true') == 'true'
        raw_terms = request.form.get('terms', '') or ''
        extra_allowed = set()
        for t in re.split(r'[\n,;]+', raw_terms):
            t = t.strip().lower()
            if t:
                extra_allowed.add(t)
        library = load_terms()                 # shared technical-terms DB
        extra_allowed |= set(library.keys())
        text = request.form.get('text', '') or ''
        token = None
        f = request.files.get('file')
        if f and f.filename:
            ext = os.path.splitext(f.filename)[1].lower()
            if ext not in ('.docx', '.txt', '.md', '.xlsx', '.xlsm'):
                return jsonify({'error': 'Unsupported file. Use .docx, .xlsx, .txt or .md, or paste text.'}), 400
            token = uuid.uuid4().hex
            path = os.path.join(TMPDIR, token + ext)
            f.save(path)
            UPLOADS[token] = {'path': path, 'filename': f.filename, 'ext': ext}
            text = extract_text_from(path, ext)
        elif text.strip():
            # store pasted text so replacement can produce a corrected copy too
            token = uuid.uuid4().hex
            path = os.path.join(TMPDIR, token + '.txt')
            with open(path, 'w', encoding='utf-8') as fh:
                fh.write(text)
            UPLOADS[token] = {'path': path, 'filename': 'pasted_text.txt', 'ext': '.txt'}
        if not text.strip():
            return jsonify({'error': 'No text provided. Paste text or upload a file.'}), 400
        result = check_text(text, mode, strict, extra_allowed)
        result['token'] = token
        result['library'] = library
        result['user'] = CURRENT_USER
        return jsonify(result)
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/apply', methods=['POST'])
def apply():
    try:
        data = request.get_json(force=True)
        token = data.get('token')
        pairs = [(p.get('find', ''), p.get('replace', '')) for p in data.get('replacements', [])]
        if token not in UPLOADS:
            return jsonify({'error': 'Session expired - run the check again.'}), 400
        pats = build_patterns(pairs)
        if not pats:
            return jsonify({'error': 'No replacements selected (tick words and set a replacement).'}), 400

        info = UPLOADS[token]
        src, ext = info['path'], info['ext']
        stem = os.path.splitext(os.path.basename(info['filename']))[0]
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        out_name = '%s_STE_%s%s' % (stem, ts, ext)
        dst = os.path.join(TMPDIR, out_name)

        if ext == '.docx':
            n = replace_in_docx(src, dst, pats)
        elif ext in ('.xlsx', '.xlsm'):
            n = replace_in_xlsx(src, dst, pats)
        else:
            n = replace_in_txt(src, dst, pats)

        try:
            resp = send_file(dst, as_attachment=True, download_name=out_name)      # Flask 2.x
        except TypeError:
            resp = send_file(dst, as_attachment=True, attachment_filename=out_name)  # Flask 1.1.x
        resp.headers['X-STE-Replacements'] = str(n)
        resp.headers['X-STE-Outfile'] = out_name
        resp.headers['Access-Control-Expose-Headers'] = 'X-STE-Replacements, X-STE-Outfile'
        return resp
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/terms', methods=['GET'])
def terms():
    return jsonify({'library': load_terms(), 'user': CURRENT_USER})

@app.route('/add_term', methods=['POST'])
def add_term():
    try:
        data = request.get_json(force=True)
        words = data.get('terms', [])
        if not words:
            return jsonify({'error': 'No words selected.'}), 400
        library, added = add_terms(words, CURRENT_USER)
        return jsonify({'ok': True, 'added': added, 'library': library, 'user': CURRENT_USER})
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

# ---------------------------------------------------------------- template
PAGE = r"""<!doctype html><html><head><meta charset="utf-8">
<title>STELOS - The ASD-STE100 Intelligence Platform</title>
<link rel="icon" href="https://www.alten.com/wp-content/uploads/2019/01/favicon-alten.png">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{--bg:#eef2f9;--surface:#fff;--surface2:#f1f5fb;--border:#c8d4e8;--accent:#1a4fad;
--accent2:#1d6fdb;--green:#059669;--red:#dc2626;--amber:#d97706;--text:#1e293b;--muted:#475569;
--mono:'IBM Plex Mono',Consolas,monospace}
body{font-family:system-ui,'Segoe UI',sans-serif;font-size:14px;background:var(--bg);color:var(--text);
min-height:100vh;display:flex;flex-direction:column}
header{background:var(--accent);color:#fff;padding:14px 28px;display:flex;align-items:center;gap:20px;
box-shadow:0 2px 8px rgba(0,0,0,.18)}
.header-logo{display:flex;align-items:center;gap:14px}
.header-logo img{height:34px;width:auto;background:#fff;border-radius:6px;padding:4px}
.brand-name{font-size:24px;font-weight:800;letter-spacing:1.5px}
.header-title h1{font-size:20px;font-weight:600}.header-title .subtitle{font-size:12px;opacity:.85;margin-top:2px}
.brand-sep{width:1px;height:34px;background:rgba(255,255,255,.35)}
.brand-by{font-size:11px;opacity:.85;margin-top:3px;letter-spacing:.5px}
main{flex:1;max-width:1040px;width:100%;margin:22px auto;padding:0 20px}
.card{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px 20px;margin-bottom:18px;
box-shadow:0 1px 3px rgba(0,0,0,.05)}
.card h2{font-size:16px;margin-bottom:12px;color:var(--accent)}
textarea{width:100%;min-height:170px;border:1px solid var(--border);border-radius:8px;padding:12px;
font-family:var(--mono);font-size:13px;resize:vertical}
.row{display:flex;gap:16px;flex-wrap:wrap;align-items:center;margin-top:12px}
label.opt{display:flex;align-items:center;gap:6px;cursor:pointer}
select,input[type=file],input.rep{padding:7px 9px;border:1px solid var(--border);border-radius:6px;background:#fff;font-size:13px}
input.rep{font-family:var(--mono);width:100%;color:var(--green)}
button{background:var(--accent2);color:#fff;border:0;padding:10px 20px;border-radius:7px;font-size:14px;
font-weight:600;cursor:pointer}button:disabled{opacity:.5;cursor:not-allowed}
button.sec{background:var(--surface2);color:var(--accent);border:1px solid var(--border)}
button.go{background:var(--green)}
.stats{display:flex;gap:14px;flex-wrap:wrap}
.stat{background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:10px 14px;min-width:110px}
.stat .n{font-size:22px;font-weight:700}.stat .l{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.4px}
.stat.bad .n{color:var(--red)}.stat.warn .n{color:var(--amber)}.stat.ok .n{color:var(--green)}
table{width:100%;border-collapse:collapse;margin-top:6px}
th,td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--border);font-size:13px;vertical-align:middle}
th{background:var(--surface2);color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.4px}
td.w{font-family:var(--mono);color:var(--red);font-weight:600}
td.chk{width:34px;text-align:center}td.repcell{width:230px}
mark.bad{background:#fde2e2;color:#991b1b;border-radius:3px;padding:0 2px;cursor:help;border-bottom:2px solid var(--red)}
mark.warnw{background:#fef3c7;color:#92400e;border-radius:3px;padding:0 2px;cursor:help}
.doc{background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:14px;line-height:1.7;
font-size:13.5px;max-height:340px;overflow:auto}
.issue{padding:8px 12px;border-left:3px solid var(--amber);background:#fffbeb;border-radius:4px;margin-bottom:8px;font-size:13px}
.issue .ctx{color:var(--muted);font-family:var(--mono);font-size:12px;margin-top:3px}
.issue.sev-error{border-left-color:var(--red);background:#fef2f2}
.issue.sev-advisory{border-left-color:var(--accent2);background:#eff6ff}
.rulebadge{display:inline-block;background:var(--accent);color:#fff;font-size:11px;font-weight:700;
border-radius:4px;padding:1px 7px;margin-right:6px}
.sevtag{font-size:10px;font-weight:700;letter-spacing:.4px;padding:1px 6px;border-radius:3px;margin-right:6px}
.sevtag.error{background:#fde2e2;color:#991b1b}.sevtag.warn{background:#fef3c7;color:#92400e}
.sevtag.advisory{background:#dbeafe;color:#1e40af}
.covsec{margin-bottom:14px}.covsec h4{font-size:13px;color:var(--muted);margin-bottom:6px;
text-transform:uppercase;letter-spacing:.4px}
.covgrid{display:grid;grid-template-columns:1fr 1fr;gap:6px}
.covitem{font-size:12.5px;padding:6px 9px;border-radius:6px;border:1px solid var(--border);
display:flex;align-items:center;gap:6px}
.covitem b{font-family:var(--mono)}
.covitem .covtag{margin-left:auto;font-size:10px;font-weight:700;padding:1px 6px;border-radius:10px}
.cov-ok{background:#f0fdf4;border-color:#bbf7d0}.cov-ok .covtag{background:#bbf7d0;color:#166534}
.cov-hit{background:#fef2f2;border-color:#fecaca}.cov-hit .covtag{background:#fecaca;color:#991b1b}
.cov-man{background:var(--surface2)}.cov-man .covtag{background:#e2e8f0;color:#475569}
@media(max-width:640px){.covgrid{grid-template-columns:1fr}}
.hidden{display:none}.err{color:var(--red);font-weight:600;margin-top:10px}
.ok-msg{color:var(--green);font-weight:600;margin-top:10px}
footer{text-align:center;padding:16px;font-size:12px;color:var(--muted)}
.tag{display:inline-block;font-size:11px;background:var(--surface2);border:1px solid var(--border);
border-radius:12px;padding:2px 10px;color:var(--muted);margin-left:6px}
.note{font-size:12px;color:var(--muted);margin:4px 0 10px}
mark.good{background:#dcfce7;color:#166534;border-radius:3px;padding:0 2px;border-bottom:2px solid var(--green)}
button.tech{background:var(--green)}
.techcell label{display:flex;align-items:center;gap:6px;font-size:12px;color:var(--muted);cursor:pointer}
.lib{margin-top:14px;border-top:1px dashed var(--border);padding-top:12px}
.lib h4{font-size:13px;color:var(--accent);margin-bottom:8px}
.chips{display:flex;flex-wrap:wrap;gap:6px}
.chip{background:#dcfce7;color:#166534;border:1px solid #bbf7d0;border-radius:14px;padding:2px 10px;
font-size:12px;font-family:var(--mono)}.chip small{color:#3f7d55;font-family:system-ui}
.rowdone{opacity:.55}
</style></head><body>
<header>
  <div class="header-logo">
    <img src="https://www.alten.com/wp-content/uploads/2019/01/favicon-alten.png" alt="ALTEN"
         onerror="this.replaceWith(Object.assign(document.createElement('span'),{textContent:'ALTEN',style:'font-weight:800;background:#fff;color:#1a4fad;padding:6px 10px;border-radius:6px'}))">
    <span class="brand-sep"></span>
    <span class="brand-name">STELOS</span>
  </div>
  <div class="header-title"><h1>The ASD-STE100 Intelligence Platform</h1>
  <p class="brand-by">An ALTEN product &middot; ASD-STE100 Simplified Technical English</p></div>
</header>
<main>
  <div class="card">
    <h2>1 &middot; Input <span class="tag">{{meta.approved_count}} approved &middot; {{meta.not_approved_count}} not approved</span></h2>
    <textarea id="text" placeholder="Paste technical text here, or upload a file below..."></textarea>
    <div class="row">
      <input type="file" id="file" accept=".docx,.xlsx,.xlsm,.txt,.md">
      <label class="opt">Writing type:
        <select id="mode">
          <option value="procedural">Procedural (max {{rules.proc_sentence_max}} words/sentence)</option>
          <option value="descriptive">Descriptive (max {{rules.desc_sentence_max}} words/sentence)</option>
        </select>
      </label>
      <label class="opt"><input type="checkbox" id="strict" checked>
        Full ASD-STE100 check (flag <b>every</b> word not in the approved dictionary)</label>
    </div>
    <p class="note">Accepted files: Word (.docx), Excel (.xlsx / .xlsm), text (.txt / .md).
    The original file is never changed &mdash; replacements are written to a downloaded copy.
    STE is a closed vocabulary: with the full check on, any word outside the approved list is
    flagged as <i>"Not in STE dictionary - verify it is an approved technical noun/verb or reword."</i>
    Turn it off to see only words that have a suggested replacement.</p>
    <details style="margin-top:6px">
      <summary style="cursor:pointer;color:var(--accent);font-size:13px">Approved technical terms (optional)</summary>
      <p class="note">Add your project's legitimate technical nouns/verbs (valve, sensor, zone, part numbers...),
      one per line or comma-separated. These will not be flagged by the full check.</p>
      <textarea id="terms" style="min-height:70px" placeholder="valve&#10;sensor&#10;zone&#10;solenoid"></textarea>
    </details>
    <div class="row">
      <button id="runBtn" onclick="run()">&#9654; Check text</button>
      <button class="sec" onclick="clearAll()">Clear</button>
      <span id="err" class="err"></span>
    </div>
  </div>

  <div id="results" class="hidden">
    <div class="card"><h2>2 &middot; Summary</h2><div class="stats" id="stats"></div></div>

    <div class="card"><h2>3 &middot; Not-approved words &amp; replacement</h2>
      <p class="note">Words with an STE alternative: tick the left box and adjust the replacement, then apply.
      Words marked <i>"Not in STE dictionary"</i> are usually your technical terms &mdash; tick
      <b>"Add as tech word"</b> and save them to the shared library so they are approved (green) for
      everyone next time. Capitalization is matched automatically.</p>
      <table><thead><tr>
        <th class="chk"><input type="checkbox" id="allChk" onclick="toggleAll(this)"></th>
        <th>Word in text</th><th>POS / rule</th><th>Count</th><th>Use instead (STE)</th><th>Replace with / tech word</th>
      </tr></thead><tbody id="wordsBody"></tbody></table>
      <div class="row">
        <button class="go" id="applyBtn" onclick="applyRepl()">&#10003; Apply ticked replacements &amp; download copy</button>
        <button class="tech" id="techBtn" onclick="addTechWords()">&#43; Add ticked words to technical library</button>
        <span id="applyMsg"></span>
      </div>
      <div class="lib">
        <h4>Shared technical-terms library <span id="libUser" class="tag"></span></h4>
        <p class="note">Saved to a shared file, visible to everyone. These words are treated as approved (green) and are never flagged.</p>
        <div class="chips" id="libChips"></div>
      </div>
    </div>

    <div class="card"><h2>4 &middot; Rule findings (Part 1 writing rules)</h2>
      <p class="note">Each finding shows the ASD-STE100 rule number. Errors break a rule outright;
      warnings and advisories need a quick check.</p>
      <div id="issues"></div></div>

    <div class="card"><h2>5 &middot; ASD-STE100 rule coverage</h2>
      <p class="note">Every rule in the standard. <b>Auto</b> = this tool checks it (issue count shown);
      <b>Manual</b> = needs your judgement (semantic rules a tool cannot decide).</p>
      <div id="coverage"></div></div>

    <div class="card"><h2>6 &middot; Annotated text</h2>
      <p class="note"><mark class="bad">red</mark> = not approved (hover for alternative) &nbsp;
      <mark class="warnw">amber</mark> = not in dictionary (strict mode)</p>
      <div class="doc" id="doc"></div>
      <div class="row"><button class="sec" onclick="dl()">&#8681; Download report (.html)</button></div>
    </div>
  </div>
</main>
<footer>STELOS &mdash; The ASD-STE100 Intelligence Platform &nbsp;&middot;&nbsp; &copy; 2026 ALTEN. All rights reserved. Confidential &ndash; Internal Use Only.</footer>
<script>
let LAST=null, TOKEN=null;
function clearAll(){document.getElementById('text').value='';document.getElementById('file').value='';
document.getElementById('results').classList.add('hidden');document.getElementById('err').textContent='';}
function esc(x){const e=document.createElement('div');e.textContent=x==null?'':x;return e.innerHTML;}

async function run(){
  const err=document.getElementById('err');err.textContent='';
  const btn=document.getElementById('runBtn');btn.disabled=true;btn.textContent='Checking...';
  const fd=new FormData();
  fd.append('text',document.getElementById('text').value);
  fd.append('mode',document.getElementById('mode').value);
  fd.append('strict',document.getElementById('strict').checked?'true':'false');
  fd.append('terms',document.getElementById('terms').value);
  const f=document.getElementById('file').files[0];if(f)fd.append('file',f);
  try{
    const r=await fetch('/check',{method:'POST',body:fd});
    const d=await r.json();
    if(!r.ok){err.textContent=d.error||'Error';return;}
    LAST=d;TOKEN=d.token;render(d);
  }catch(e){err.textContent=e.message;}
  finally{btn.disabled=false;btn.innerHTML='&#9654; Check text';}
}

function render(d){
  document.getElementById('results').classList.remove('hidden');
  document.getElementById('applyMsg').textContent='';
  const s=d.summary;
  document.getElementById('stats').innerHTML=`
    <div class="stat"><div class="n">${s.words}</div><div class="l">Words</div></div>
    <div class="stat"><div class="n">${s.sentences}</div><div class="l">Sentences</div></div>
    <div class="stat ${s.unapproved_total?'bad':'ok'}"><div class="n">${s.unapproved_total}</div><div class="l">Not-approved uses</div></div>
    <div class="stat warn"><div class="n">${s.length_issues}</div><div class="l">Long sentences</div></div>
    <div class="stat warn"><div class="n">${s.paragraph_issues}</div><div class="l">Long paragraphs</div></div>
    <div class="stat ${s.rule_issues?'warn':'ok'}"><div class="n">${s.rule_issues}</div><div class="l">Rule findings</div></div>`;
  const wb=document.getElementById('wordsBody');
  if(d.flagged.length===0){wb.innerHTML='<tr><td colspan="6" style="color:var(--green)">No not-approved words found. &#10003;</td></tr>';}
  else{wb.innerHTML=d.flagged.map((f,i)=>{
     const lastcell = f.has_alt
       ? `<td class="repcell"><input class="rep" data-i="${i}" value="${esc(f.replacement)}"
            placeholder="type replacement..."></td>`
       : `<td class="techcell"><label><input type="checkbox" class="techchk" data-word="${esc(f.word)}">
            &#43; Add as tech word</label></td>`;
     return `<tr>
       <td class="chk">${f.has_alt?`<input type="checkbox" class="rowchk" data-i="${i}" checked>`:''}</td>
       <td class="w">${esc(f.word)}</td><td>${esc(f.pos)}</td><td>${f.count}</td>
       <td style="color:var(--green);font-family:var(--mono)">${esc(f.alt)}</td>
       ${lastcell}</tr>`;}).join('');}
  renderLibrary(d.library, d.user);
  const iss=document.getElementById('issues');
  const rf=d.rule_findings||[];
  if(rf.length===0){iss.innerHTML='<p style="color:var(--green)">No rule findings. &#10003;</p>';}
  else{
    const order={error:0,warn:1,advisory:2};
    const sevlabel={error:'ERROR',warn:'WARNING',advisory:'ADVISORY'};
    const sorted=rf.slice().sort((a,b)=>(order[a.severity]-order[b.severity]));
    iss.innerHTML=sorted.map(i=>`<div class="issue sev-${i.severity}">
       <span class="rulebadge">Rule ${esc(i.rule)}</span>
       <span class="sevtag ${i.severity}">${sevlabel[i.severity]||''}</span>
       ${esc(i.msg)}<div class="ctx">${esc(i.context)}</div></div>`).join('');
  }
  const cov=document.getElementById('coverage');
  cov.innerHTML=(d.coverage||[]).map(sec=>`
    <div class="covsec"><h4>${esc(sec.section)}</h4>
    <div class="covgrid">${sec.rules.map(r=>{
      const auto=r.kind==='auto';
      const cls=auto?(r.count?'cov-hit':'cov-ok'):'cov-man';
      const tag=auto?(r.count?(r.count+' issue'+(r.count>1?'s':'')):'clear'):'manual';
      return `<div class="covitem ${cls}"><b>${esc(r.num)}</b> ${esc(r.title)}
              <span class="covtag">${tag}</span></div>`;}).join('')}</div></div>`).join('');
  document.getElementById('doc').innerHTML=d.highlighted;
}

function toggleAll(cb){document.querySelectorAll('.rowchk').forEach(c=>c.checked=cb.checked);}

function renderLibrary(lib, user){
  document.getElementById('libUser').textContent = user ? ('you: '+user) : '';
  const box=document.getElementById('libChips');
  const keys=Object.keys(lib||{}).sort();
  if(keys.length===0){box.innerHTML='<span class="note">No technical terms saved yet.</span>';return;}
  box.innerHTML=keys.map(k=>`<span class="chip">${esc(k)} <small>&middot; ${esc((lib[k]&&lib[k].by)||'')}</small></span>`).join('');
}

async function loadLibrary(){
  try{const r=await fetch('/terms');const d=await r.json();renderLibrary(d.library,d.user);}catch(e){}
}

async function addTechWords(){
  const msg=document.getElementById('applyMsg');msg.className='';msg.textContent='';
  const words=[...document.querySelectorAll('.techchk')].filter(c=>c.checked).map(c=>c.dataset.word);
  if(words.length===0){msg.className='err';msg.textContent='Tick at least one "Add as tech word" box.';return;}
  const btn=document.getElementById('techBtn');btn.disabled=true;btn.textContent='Saving...';
  try{
    const r=await fetch('/add_term',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({terms:words})});
    const d=await r.json();
    if(!r.ok){msg.className='err';msg.textContent=d.error||'Error';return;}
    renderLibrary(d.library,d.user);
    msg.className='ok-msg';msg.textContent='Saved '+d.added.length+' technical term(s) to the shared library. Re-checking...';
    // grey the added rows immediately, then re-check so they turn green / drop off
    document.querySelectorAll('.techchk').forEach(c=>{if(c.checked)c.closest('tr').classList.add('rowdone');});
    setTimeout(run, 500);
  }catch(e){msg.className='err';msg.textContent=e.message;}
  finally{btn.disabled=false;btn.innerHTML='&#43; Add ticked words to technical library';}
}

async function applyRepl(){
  const msg=document.getElementById('applyMsg');msg.className='';msg.textContent='';
  if(!TOKEN){msg.className='err';msg.textContent='Run a check first.';return;}
  const reps=[];
  document.querySelectorAll('.rowchk').forEach(chk=>{
    if(chk.checked){
      const i=chk.dataset.i;const f=LAST.flagged[i];
      const rep=document.querySelector('.rep[data-i="'+i+'"]').value.trim();
      if(rep) reps.push({find:f.word,replace:rep});
    }
  });
  if(reps.length===0){msg.className='err';msg.textContent='Tick at least one word that has a replacement.';return;}
  const btn=document.getElementById('applyBtn');btn.disabled=true;btn.textContent='Working...';
  try{
    const r=await fetch('/apply',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({token:TOKEN,replacements:reps})});
    if(!r.ok){const d=await r.json();msg.className='err';msg.textContent=d.error||'Error';return;}
    const n=r.headers.get('X-STE-Replacements');const name=r.headers.get('X-STE-Outfile')||'corrected';
    const blob=await r.blob();const a=document.createElement('a');
    a.href=URL.createObjectURL(blob);a.download=name;a.click();
    msg.className='ok-msg';msg.textContent='Done - '+n+' replacement(s) written to a copy: '+name;
  }catch(e){msg.className='err';msg.textContent=e.message;}
  finally{btn.disabled=false;btn.innerHTML='&#10003; Apply ticked replacements &amp; download copy';}
}

function dl(){
  if(!LAST)return;
  const doc=document.getElementById('doc').innerHTML;const s=LAST.summary;
  const rows=LAST.flagged.map(f=>`<tr><td>${esc(f.word)}</td><td>${esc(f.pos)}</td><td>${f.count}</td><td>${esc(f.alt)}</td></tr>`).join('');
  const isr=(LAST.rule_findings||[]).map(i=>`<li><b>Rule ${esc(i.rule)}</b> [${esc(i.severity)}] ${esc(i.msg)} <i>${esc(i.context)}</i></li>`).join('');
  const cov=(LAST.coverage||[]).map(sec=>`<h3>${esc(sec.section)}</h3><table><tr><th>Rule</th><th>Title</th><th>Coverage</th></tr>`+
    sec.rules.map(r=>`<tr><td>${esc(r.num)}</td><td>${esc(r.title)}</td><td>${r.kind==='auto'?(r.count?('AUTO - '+r.count+' issue(s)'):'AUTO - clear'):'Manual review'}</td></tr>`).join('')+`</table>`).join('');
  const h=`<!doctype html><meta charset=utf-8><title>STELOS - ASD-STE100 report</title>
  <link rel="icon" href="https://www.alten.com/wp-content/uploads/2019/01/favicon-alten.png">
  <style>body{font-family:Segoe UI,sans-serif;max-width:900px;margin:30px auto;color:#1e293b}
  h1{color:#1a4fad}h3{color:#475569;margin-top:14px}mark.bad{background:#fde2e2;border-bottom:2px solid #dc2626}mark.warnw{background:#fef3c7}
  table{border-collapse:collapse;width:100%;margin-bottom:8px}td,th{border:1px solid #c8d4e8;padding:6px;text-align:left;font-size:13px}
  .doc{background:#f1f5fb;padding:14px;border-radius:8px;line-height:1.7}</style>
  <h1>STELOS &ndash; ASD-STE100 report</h1><p style="color:#475569;margin-top:-6px">The ASD-STE100 Intelligence Platform &middot; An ALTEN product</p>
  <p>${new Date().toLocaleString()} &middot; ${s.mode} mode (limit ${s.sentence_limit} words)</p>
  <p><b>${s.words}</b> words, <b>${s.sentences}</b> sentences, <b>${s.unapproved_total}</b> not-approved uses,
  <b>${s.length_issues}</b> long sentences, <b>${s.paragraph_issues}</b> long paragraphs, <b>${s.rule_issues}</b> rule findings.</p>
  <h2>Not-approved words</h2><table><tr><th>Word / rule</th><th>Tag</th><th>Count</th><th>Use instead</th></tr>${rows}</table>
  <h2>Rule findings</h2><ul>${isr||'<li>None</li>'}</ul>
  <h2>ASD-STE100 rule coverage</h2>${cov}
  <h2>Annotated text</h2><div class="doc">${doc}</div>
  <hr><small>STELOS &middot; &copy; 2026 ALTEN. Confidential - Internal Use Only.</small>`;
  const b=new Blob([h],{type:'text/html'});const a=document.createElement('a');
  a.href=URL.createObjectURL(b);a.download='STE_report_'+Date.now()+'.html';a.click();
}
loadLibrary();
</script>
</body></html>"""

def open_browser():
    time.sleep(1.2)
    webbrowser.open('http://127.0.0.1:%d' % PORT)

if __name__ == '__main__':
    print('STELOS - The ASD-STE100 Intelligence Platform (ALTEN) | %s | %d approved / %d not-approved' %
          (META['source'], META['approved_count'], META['not_approved_count']))
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host='127.0.0.1', port=PORT, debug=False, use_reloader=False)
