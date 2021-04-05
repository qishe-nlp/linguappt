from linguappt.en._vocab_meta import _EnglishVocabMeta

def test_vocabmeta():
  pos = _EnglishVocabMeta.get_pos("v.vi.")
  assert pos == "verb"
