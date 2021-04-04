# Usage

### Install from pip3

```
pip3 install --verbose linguappt 
```

### Execute usage

* Validate ppt template
```
pptx_validate --pptx [pptx file]
```

* Convert vocabulary csv file into ppt file
```
lingua_vocabppt --sourcecsv [vocab csv file] --lang [language] --title [title shown in ppt] --destpptx [pptx file]
```

* Convert phrase csv file into ppt file
```
lingua_vocabppt --sourcecsv [phrase csv file] --lang [language] --title [title shown in ppt] --destpptx [pptx file]
```


* Convert ppt into pdf
```
lingua_pptx2pdf --sourcepptx [pptx file] --destdir [dest directory storing pdf and images]
```

### Package usage
```
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

### Test
```
poetry run pytest -rP
```
which run tests under `tests/*`


### Execute
```
poetry run pptx_validate --help
poetry run vocab_csv2ppt --help
poetry run ppt2pdf --help
poetry run vocab_csv2pptpdf --help
```

### Create sphinx docs
```
poetry shell
cd apidocs
sphinx-apidoc -f -o source ../subtitlecore
make html
python -m http.server -d build/html
```

### Host docs on github pages
```
cp -rf apidocs/build/html/* docs/
```

### Build
* Change `version` in `pyproject.toml` and `linguappt/__init__.py`
* Build python package by `poetry build`

### Git commit and push

### Publish from local dev env
* Set pypi test environment variables in poetry, refer to [poetry doc](https://python-poetry.org/docs/repositories/)
* Publish to pypi test by `poetry publish -r test`

### Publish through CI 
* Github action build and publish package to [test pypi repo](https://test.pypi.org/)

```
git tag [x.x.x]
git push origin master
```

* Manually publish to [pypi repo](https://pypi.org/) through [github action](https://github.com/qishe-nlp/linguappt/actions/workflows/pypi.yml)

