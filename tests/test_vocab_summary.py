from linguappt.en.vocab_summary import EnglishVocabMeta

def test_vocabmeta():
  pos = EnglishVocabMeta.get_pos("vi.")
  assert pos == "verb"
