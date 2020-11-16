# Usage

### Install from pip3

```
pip3 install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple --verbose linguappt 
```

### Execute usage

* Validate ppt template
```
pptx_validate --pptx [pptx file]
```

* Convert vocabulary csv file into ppt file
```
vocab_csv2ppt --sourcecsv [vocab csv file] --title [title shown in ppt] --destpptx [pptx file]
```

* Convert ppt into pdf
```
ppt2pdf --sourcepptx [pptx file] --destdir [dest directory storing pdf and images]
```

Generate pptx from csv which contains vocabulary information 
```
vocab_csv2pptpdf --sourcecsv [vocab csv file] --name [name of ppt pdf img] --pptxdir [pptx directory] --pdfdir [pdf directory] --imgdir [image directory]
```

### Package usage
```
from linguappt import SpanishVocabPPT, Pdf
import os

def vocab_csv2pptpdf(sourcecsv, name, pptxdir, pdfdir, imgdir):

  pptx_pro_dir = pptxdir + '/pro/' 
  pptx_water_dir = pptxdir + '/water/' 

  if not os.path.isdir(pptxdir):
    os.mkdir(pptxdir) 

  if not os.path.isdir(pptx_pro_dir):
    os.mkdir(pptx_pro_dir) 

  if not os.path.isdir(pptx_water_dir):
    os.mkdir(pptx_water_dir) 

  vp = SpanishVocabPPT(sourcecsv, "歧舌AI备课助教")
  pptx = pptx_pro_dir + name +'.pptx'
  vp.convert_to_ppt(pptx)

  vp = SpanishVocabPPT(sourcecsv, "歧舌AI备课助教", "watermark")
  watermark_pptx = pptx_water_dir + name + ".pptx"
  vp.convert_to_ppt(watermark_pptx)

  if not os.path.isdir(pdfdir):
    os.mkdir(pdfdir) 

  pdf_pro_dir = pdfdir + '/pro/'
  pdf_water_dir = pdfdir + '/water/'
  pdf = Pdf(pptx, pdf_pro_dir)
  watermark_pdf = Pdf(watermark_pptx, pdf_water_dir)

  images_len = pdf.save_as_images(0, 6, imgdir)
```

# Development

### Clone project
```
git clone https://github.com/qishe-nlp/linguappt.git
```

### Install [poetry](https://python-poetry.org/docs/)

### Install dependencies
```
poetry update
```

### Execute
```
poetry run pptx_validate --help
poetry run vocab_csv2ppt --help
poetry run ppt2pdf --help
poetry run vocab_csv2pptpdf --help
```

### Build
* Change `version` in `pyproject.toml` and `linguappt/__init__.py`
* Build python package by `poetry build`

### Publish
* Set pypi test environment variables in poetry, refer to [poetry doc](https://python-poetry.org/docs/repositories/)
* Publish to pypi test by `poetry publish -r test`


# TODO

### Test and Issue
* `tests/*`

### Github action to publish package
* pypi test repo
* pypi repo
