import pytest
from macscript import asvocab
from macscript.word import SDEF, sd, sdfind


def test_word_vocabulary():
    assert 'class document' in sd('document')
    assert any(kind == 'property' and name.startswith('document.') for kind, name, _ in sdfind('name'))
    assert sd('save_as') == asvocab.sd('save as', SDEF)
    with pytest.raises(KeyError): sd('not a Word term')


def test_other_app_vocabulary(tmp_path):
    path = tmp_path/'example.sdef'
    path.write_text('<dictionary><suite><command name="greet" description="Say hello">'
        '<parameter name="name" type="text" optional="yes"/></command></suite></dictionary>')
    assert 'name: text (optional)' in asvocab.sd('greet', path)
    assert asvocab.sdfind('hello', path) == [('command', 'greet', 'Say hello')]
