# macscript

Script macOS applications from Python, using appscript. Start with Microsoft Word; use the vocabulary helpers to explore any scriptable app.

Requires macOS. Word automation also requires Microsoft Word. macOS may ask you to allow Automation access; dialog controls use Accessibility access, and screenshots need Screen Recording access.

## Word

```python
from macscript.word import new_doc, set_text, doc_text, save_docx, close_doc

doc = new_doc()
set_text(doc, 'Hello from Python!')
print(doc_text(doc))
doc = save_docx(doc, 'hello.docx')
close_doc(doc)
```

`open_doc(path)` opens an existing document. `save_docx` returns a fresh document reference because saving under a new name makes the old reference stale. `save_pdf(doc, path)` exports a PDF. `close_doc` discards unsaved changes unless passed `save=True`. Body text returned by `doc_text` uses LF line endings.

`word()` returns the appscript application proxy for commands not wrapped here. Document operations accept a `timeout` in seconds; the default is 15.

### Screenshots and dialogs

- `win_pic(path)` captures Word's front document window on the current Space. An occluded window can still be captured.
- `win_id()` returns its window ID; `win2png(window_id, path)` captures a window by ID.
- `dlg_read()` returns the front dialog's button names and message, or `None`.
- `dlg_click(label)` clicks a named button.
- `dismiss_dlgs()` dismisses dialogs and returns their messages. It chooses OK, then No, then the first available button; use it only when that behavior is appropriate.

Dialog queries default to a three-second timeout. They use System Events so they can work while Word's own event queue is blocked by a modal dialog.

## Discover scripting commands

Word's scripting dictionary is bundled, so its documentation can be searched without launching Word:

```python
from macscript.word import sd, sdfind

print(sd('document'))
print(sd('save_as'))
sdfind('bookmark')
```

`sd` shows a command, class, or enumeration with types and descriptions. Underscores and spaces both work in names. `sdfind` searches names and descriptions with a case-insensitive regex and returns `(kind, name, description)` rows.

The generic helpers work with other appscript apps too:

```python
from appscript import app
from macscript.asvocab import vocab, props, sd, sdfind

finder = app('Finder')
vocab(finder, 'selection|window')
props(finder, 'name')
sd('application', path='Finder.sdef')
```

`vocab` searches an app's live terminology. `props` reads a reference's properties as a dictionary of names and truncated representations, with optional regex filtering. Generic `sd` and `sdfind` require an explicit `.sdef` path; only the Word module supplies a Word default.

## Development

Install the checkout and its development dependencies:

```bash
pip install -e '.[dev]'
python -m pytest -q
```

The default tests exercise scripting documentation without launching apps. There is no dependency on `mdhtml2docx` or `oxml`; document validation and converter acceptance checks belong to their callers. App-specific behavior lives in app modules, with reusable scripting exploration in `asvocab`.
