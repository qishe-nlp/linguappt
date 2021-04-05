from linguappt.vocab_ppt import VocabPPT
import pytest

def test_phraseppt():
  with pytest.raises(TypeError):
    VocabPPT("./test.csv:")
