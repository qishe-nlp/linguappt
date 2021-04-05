from linguappt.phrase_ppt import PhrasePPT
import pytest

def test_phraseppt():
  with pytest.raises(TypeError):
    PhrasePPT("./test.csv:")
