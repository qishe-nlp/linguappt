from linguappt.lib import readCSV, pptx2pdf, pdf2images

def test_readCSV(capsys):
  none_existed_file = "xxxx.csv"
  existed_file = "en_phrase.forpptx.csv"

  content = readCSV(none_existed_file)
  assert len(content) == 0
  captured = capsys.readouterr()
  assert captured.err == "xxxx.csv DOES NOT exist!!!\n"
  
  content = readCSV(existed_file)
  assert len(content) > 0


def test_pdf2images():
  pass

def test_pptx2pdf():
  pass
