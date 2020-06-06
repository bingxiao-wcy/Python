'''
from nose.tools import*
from ex49 import scan

def test_peek():
    words = [('noun','apple'),('stop','the'),('verb','kill')]
    assert_equal(scan.peek(words),'noun')

def test_parse_verb():
    words = [('noun','apple'),('stop','the'),('verb','kill')]
    assert_equal(scan.parse_verb(words), ('verb','kill'))

def test_match():
    words = [('noun','apple'),('stop','the'),('verb','kill')]
    assert_equal(scan.match(words, 'noun'),('noun','apple'))
    '''
from nose.tools import *
from ex49 import scan
def test_peek():
    words = [('noun', 'apple'), ('stop', 'the'), ('verb', 'kill')]
    assert_equal(scan.peek(words), 'noun')
    assert_equal(scan.peek([]), None)
def test_match():
    assert_equal(scan.match([], 'noun'), None)
    words = [('noun','apple'),('stop','the'), ('verb','kill')]
    assert_equal(scan.match(words, 'noun'), ('noun', 'apple'))
    assert_equal(scan.match(words, 'verb'), None)
# def test_skip():
def test_parse_verb():
    words=[('stop', 'a'), ('stop', 'the'), ('verb', 'kill'), ('noun', 'apple')]
    assert_equal(scan.parse_verb(words), ('verb', 'kill'))
    words_error=[('stop', 'a'), ('stop', 'the'), ('noun', 'apple'), ('verb', 'kill')]
    assert_raises(scan.ParserError, scan.parse_verb, words_error)
def test_parse_object():
    words = [('noun', 'apple'), ('verb', 'kill')]
    assert_equal(scan.parse_object(words), ('noun', 'apple'))
    words2 = [('direction', 'east'), ('verb', 'kill')]
    assert_equal(scan.parse_object(words2), ('direction', 'east'))
    words_error = [('stop', 'a'), ('verb', 'run'), ('verb', 'kill')]
    assert_raises(scan.ParserError, scan.parse_object, words_error)
def test_parse_sentence():
    words = [('stop', 'a'), ('noun', 'apple'), ('verb', 'kill'), ('direction', 'east')]
    result = scan.parse_sentence(words)
    assert_equal(result.subject, 'apple')
    assert_equal(result.verb, 'kill')
    assert_equal(result.object, 'east')
