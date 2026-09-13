"Script Microsoft Word: documents, screenshots, dialogs, and scripting vocabulary."
from functools import cache
from importlib.resources import files
from pathlib import Path
from appscript import app, k, mactypes
import Quartz, time
from Foundation import NSURL
from . import asvocab


__all__ = ['TIMEOUT', 'UI_TIMEOUT', 'word', 'sd', 'sdfind', 'new_doc', 'open_doc', 'set_text', 'doc_text', 'save_docx',
    'save_pdf', 'close_doc', 'win_id', 'win2png', 'win_pic', 'word_proc', 'dlg_read', 'dlg_click', 'dismiss_dlgs']

TIMEOUT = 15  # default seconds for every AppleEvent; Word blocked on a modal is the failure this bounds
UI_TIMEOUT = 3  # System Events UI queries answer near-instantly or never; this bounds the never
SDEF = files('macscript')/'word.sdef'

def sd(name):
    "Full scripting documentation for a Word command, class, or enumeration"
    return asvocab.sd(name, SDEF)

def sdfind(pat, maxlen=80):
    "Search Word's scripting names and descriptions by regex"
    return asvocab.sdfind(pat, SDEF, maxlen)

@cache
def word():
    "Word application proxy; `terms='sdef'` is required since the default terminology fetch returns only built-ins"
    return app('Microsoft Word', terms='sdef')

def _val(x):
    "Normalize Word return conventions: `missing value` to None, CR line endings to LF"
    if x == k.missing_value: return None
    if isinstance(x, str): return x.replace('\r', '\n')
    return x

def new_doc(timeout=None):
    "Create an empty document, returning its reference"
    return word().make(new=k.document, timeout=timeout or TIMEOUT)

def open_doc(path, timeout=None):
    "Open `path` in Word, returning the document reference"
    p = Path(path).expanduser().resolve()
    word().open(mactypes.Alias(str(p)), timeout=timeout or TIMEOUT)
    return word().documents[p.name]

def set_text(doc, text, timeout=None):
    "Replace `doc`'s body text"
    doc.text_object.content.set(text, timeout=timeout or TIMEOUT)

def doc_text(doc, timeout=None):
    "Body text of `doc`, normalized"
    return _val(doc.text_object.content.get(timeout=timeout or TIMEOUT))

def save_docx(doc, path, timeout=None):
    "Save `doc` as modern docx to `path`; renames the doc, so returns its fresh reference (old refs go stale)"
    p = Path(path).expanduser().resolve()
    doc.save_as(file_name=str(p), file_format=k.format_document_default, timeout=timeout or TIMEOUT)
    return word().documents[p.name]

def save_pdf(doc, path, timeout=None):
    "Export `doc` as PDF to `path`, returning it"
    p = Path(path).expanduser().resolve()
    doc.save_as(file_name=str(p), file_format=k.format_PDF, timeout=timeout or TIMEOUT)
    return p

def close_doc(doc, save=False, timeout=None):
    "Close `doc`, without saving unless asked"
    doc.close(saving=k.yes if save else k.no, timeout=timeout or TIMEOUT)

def win_id():
    "CGWindow id of Word's frontmost document window on the current Space"
    wins = Quartz.CGWindowListCopyWindowInfo(Quartz.kCGWindowListOptionOnScreenOnly|Quartz.kCGWindowListExcludeDesktopElements, Quartz.kCGNullWindowID)
    ids = [w['kCGWindowNumber'] for w in wins if w['kCGWindowOwnerName']=='Microsoft Word' and w['kCGWindowLayer']==0]
    if not ids: raise ValueError('No Microsoft Word window on screen')
    return ids[0]

def win2png(wid, path):
    "Capture the window with CGWindow id `wid` (occluded is fine; off-Space windows are not capturable) to png at `path`, returning it"
    # CGWindowListCreateImage is deprecated (successor: ScreenCaptureKit, async/delegate API) but works through macOS 26
    img = Quartz.CGWindowListCreateImage(Quartz.CGRectNull, Quartz.kCGWindowListOptionIncludingWindow, wid, Quartz.kCGWindowImageBoundsIgnoreFraming)
    if img is None or Quartz.CGImageGetWidth(img) <= 1: raise ValueError(f'could not capture window {wid} (gone, or Screen Recording not permitted)')
    p = Path(path).expanduser().resolve()
    dst = Quartz.CGImageDestinationCreateWithURL(NSURL.fileURLWithPath_(str(p)), 'public.png', 1, None)
    Quartz.CGImageDestinationAddImage(dst, img, None)
    if not Quartz.CGImageDestinationFinalize(dst): raise ValueError(f'could not write png to {p}')
    return p



def win_pic(path):
    "Capture Word's document window (occluded is fine) to png at `path`, returning it"
    return win2png(win_id(), path)

@cache
def word_proc():
    "System Events proxy for Word's process, for driving dialogs while Word's own event queue is blocked"
    return app('System Events').processes['Microsoft Word']

def dlg_read(timeout=None):
    "Front dialog window as (button names, message text), or None if the front window isn't a dialog"
    to = timeout or UI_TIMEOUT
    wp = word_proc()
    names = wp.windows.name.get(timeout=to)
    if not names or names[0] not in ('', k.missing_value): return None
    w = wp.windows[1]
    btns = w.buttons.name.get(timeout=to)
    txt = '\n'.join(_val(t) for t in w.static_texts.value.get(timeout=to) if t and t != 'Microsoft Word')
    return btns, txt

def dlg_click(label, timeout=None):
    "Click button `label` on the front dialog window"
    word_proc().windows[1].buttons[label].click(timeout=timeout or UI_TIMEOUT)

def dismiss_dlgs(wait=1):
    "Dismiss Word's dialogs (OK/No, i.e. declining recovery), polling `wait` secs for chained ones; returns their message texts"
    msgs = []
    while True:
        end = time.time() + wait
        while not (d := dlg_read()) and time.time() < end: time.sleep(0.1)
        if not d: return msgs
        btns, txt = d
        msgs.append(txt)
        dlg_click('OK' if 'OK' in btns else 'No' if 'No' in btns else btns[0])
