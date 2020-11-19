from linguappt.ppt import PPT
from linguappt.filereader import readCSV
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import os
import math
import json

class EnglishVocabMeta:

  pos_info = {
    'noun': ('名词', 'noun', ['n.']),
    'adj': ('形容词', 'adjective', ['adj.', 'a.']),
    'verb': ('动词', 'verb', ['v.', 'vt.vi.', 'vi.', 'vt.', 'aux.', 'vi.vt.']),
    'adv': ('副词', 'adverb', ['adv.']),
    'pron': ('代词', 'pronoun', ['pron.']),
    'prep': ('前置词', '', ['prep.']),
    'other': ('其他', 'others', [])
  }

  tense_info = {
    'imperativo_afirmativo': "命令式-肯定",
    'imperativo_negativo': "命令式-否定",
    'indicativo-pretérito': "陈述式-过去时",
    'indicativo-presente': "陈述式-现在时",
    'subjuntivo-presente': "虚拟式-现在时",
    'indicativo-futuro': "陈述式-将来时",
    'subjuntivo-futuro': "虚拟式-将来时",
    'indicativo-imperfecto': "陈述式-未完成时",
    'subjuntivo-imperfecto': "虚拟式-未完成时",
    'indicativo-condicional': "条件式",
    'participio': "过去分词",
    'gerundio': "现在分词"
  }

class EnglishVocabPPT(PPT):

  template_dir = os.path.dirname(__file__)
  templates = {
    "classic": os.path.join(template_dir, 'templates/vocab_english_classic.pptx'),
    "watermark": os.path.join(template_dir, 'templates/vocab_english_watermark.pptx')
  }
  lang = 'en'

  def __init__(self, sourcefile, title="", genre="classic"):
    """
    word_distribution is a list of dict
    e.g,
    word_distribution = [
      {word: 'Hello', pos: 'INTJ', num: 2},
      {word: 'good', pos: 'ADJ', num: 20},
      {word: 'then', pos: 'ADV', num: 4},
    ]
    """

    self.template = EnglishVocabPPT.templates[genre]

    self.keys = ['num', 'word', 'pos', 'meaning', 'dict_pos', 'from', 'extension', 'variations', 'examples']
    self.content = self.assert_content(sourcefile)

    self.divide_by_pos()
    self.cal_displayed_word_distribution()
    self.cal_distribution()

    self.prs = Presentation(self.template)
    self.title = title
    
  def assert_content(self, sourcefile):
    """
      Ensure the content has keys defined in ppt class
    """
    content = readCSV(sourcefile) 
    for e in content:
      keys = e.keys()
      assert(len(keys) == len(self.keys))
      for k in keys:
        assert(k in self.keys)
    return content

  def divide_by_pos(self):
    """
      Store words in self.word_distribution as dict, whose key is dict_pos and value is a list of words
      dict_pos is defined in x2cdict package
    """
    self.word_distribution  = {}

    for e in self.content:
      key = e["dict_pos"]
      if key not in self.word_distribution:
        self.word_distribution[key] = []
      self.word_distribution[key].append(e)

  def cal_displayed_word_distribution(self):
    """
      Bridge dict_pos and pos defined in EnglishVocabMeta
      Store words in self.displayed_word_distribution as dict, which are used to presented in ppt.
      Key in self.displayed_word_distribution is pos defined EnglishVocabMeta and value is a list of words
    """
    self.displayed_word_distribution = {}
    for key in EnglishVocabMeta.pos_info.keys():
      poses = EnglishVocabMeta.pos_info[key][2] 
      self.displayed_word_distribution[key] = []
      for pos in poses:
        if pos in self.word_distribution.keys():
          self.displayed_word_distribution[key].extend(self.word_distribution[pos])


  def cal_distribution(self):
    """
      Calculate the number of words according to pos defined in EnglishVocabMeta
    """
    self.distribution = [{"pos": pos, "name": EnglishVocabMeta.pos_info[pos][1], "num": len(ws)} for pos, ws in self.displayed_word_distribution.items()]

  def create_home(self):
    layout = self.prs.slide_layouts.get_by_name("Title and subtitle for chinese")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders
 
    title = holders[10] 
    title.text_frame.text = self.title
    subtitle = holders[11]
    subtitle.text_frame.text = '词汇总结'

  def create_statistic(self):
    layout = self.prs.slide_layouts.get_by_name("Word count")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders

    for index, info in enumerate(self.distribution[:3]):
      pos = holders[11+2*index]
      num = holders[10+2*index]
      pos.text_frame.text = info["name"].upper() 
      num.text_frame.text = str(info["num"])


  def create_vocab_title(self, index, title_content, subtitle_content):
    layout = self.prs.slide_layouts.get_by_name('Title '+str(index))
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders

    title = holders[10]
    title.text_frame.text = title_content
    subtitle = holders[11]
    subtitle.text_frame.text = subtitle_content
    note = slide.notes_slide
    note.notes_text_frame.text = title_content.lower()


  def create_noun_word(self, v):
    layout = self.prs.slide_layouts.get_by_name("Noun vocab")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders

    noun, meaning = holders[11], holders[12]
    noun.text_frame.text = v["word"]
    ms = v["meaning"].split(",")
    if len(ms) > 4:
      ms = ms[:4]
    meaning.text_frame.text = "\n".join(ms) 

    if v["extension"] != "":
      extension = json.loads(v["extension"])
      single, plural = holders[13], holders[14]
      single.text_frame.text = v["word"]
      plural.text_frame.text = extension["s"]

    if v["examples"] != "":
      examples = json.loads(v["examples"])
      if len(examples) == 2:
        original, translated = holders[15], holders[16]
        ex = examples[0]
        original.text_frame.text = ex["original"]
        translated.text_frame.text = ex["translated"]

        original, translated = holders[17], holders[18]
        ex = examples[1]
        original.text_frame.text = ex["original"]
        translated.text_frame.text = ex["translated"]

    note = slide.notes_slide
    note.notes_text_frame.text = v["word"]


  def create_noun_mpl_word(self, v):
    layout = self.prs.slide_layouts.get_by_name("Noun m vocab")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders

    noun, meaning = holders[10], holders[11]
    noun.text_frame.text = v["word"]
    ms = v["meaning"].split(",")
    if len(ms) > 4:
      ms = ms[:4]
    meaning.text_frame.text = "\n".join(ms) 

    s_def, s = holders[12], holders[13]
    s_def.text_frame.text = "los"
    s.text_frame.text = v["word"]

    s_undef, s = holders[14], holders[15]
    s_undef.text_frame.text = "unos"
    s.text_frame.text = v["word"]

    if v["extension"] != "":
      extension = json.loads(v["extension"].replace("\'", "\""))

      pl_def, pl = holders[16], holders[17]
      pl_def.text_frame.text = "el"
      pl.text_frame.text = extension["m"]

      pl_undef, pl = holders[18], holders[19]
      pl_undef.text_frame.text = "uno"
      pl.text_frame.text = extension["m"]

    note = slide.notes_slide
    note.notes_text_frame.text = v["word"]

  def create_adj_word(self, v):
    layout = self.prs.slide_layouts.get_by_name("Adj vocab")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders

    adj, meaning = holders[11], holders[12]
    adj.text_frame.text = v["word"]
    ms = v["meaning"].split(",")
    if len(ms) > 4:
      ms = ms[:4]

    meaning.text_frame.text = "\n".join(ms)

    much, more, most = holders[13], holders[14], holders[15] 
    if v["extension"] != "":
      extension = json.loads(v["extension"])

      much.text_frame.text = extension['original'] 
      more.text_frame.text = extension['comparative']
      most.text_frame.text = extension['superlative']

    note = slide.notes_slide
    note.notes_text_frame.text = v["word"]

  def create_default_word(self, v):
    layout = self.prs.slide_layouts.get_by_name("Common layout")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders
   
    pos = holders[12]
    word = holders[13]
    meaning = holders[14]
    pos.text_frame.text = v["dict_pos"]
    word.text_frame.text = v["word"] 
    ms = v["meaning"].split(",")
    if len(ms) > 4:
      ms = ms[:4]

    meaning.text_frame.text = "\n".join(ms)
 
    if v["examples"] != "":
      examples = json.loads(v["examples"])
      if len(examples) == 2:
        original, translated = holders[15], holders[16]
        ex = examples[0]
        original.text_frame.text = ex["original"]
        translated.text_frame.text = ex["translated"]

        original, translated = holders[17], holders[18]
        ex = examples[1]
        original.text_frame.text = ex["original"]
        translated.text_frame.text = ex["translated"]

    note = slide.notes_slide
    note.notes_text_frame.text = v["word"]

  def create_single_tiempo_verb_word(self, v):
    layout = self.prs.slide_layouts.get_by_name("Verb single tiempo")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders
    
    variations = json.loads(v["variations"].replace("\'", "\""))

    origin, word, meaning = holders[10], holders[11], holders[12]
    origin.text_frame.text = variations["origin"]
    word.text_frame.text = v["word"]

    ms = v["meaning"].split(",")
    if len(ms) > 4:
      ms = ms[:4]

    meaning.text_frame.text = "\n".join(ms)
 

    sign = variations["formats"][0]
    tense = sign["tense"]
    person = sign["person"]

    extension = json.loads(v["extension"].replace("\'", "\""))[tense]

    holders[13].text_frame.text = extension["yo"] if extension["yo"] != "" else " "
    holders[14].text_frame.text = extension["tú"]
    holders[15].text_frame.text = extension["él/ella/Usted"]
    holders[16].text_frame.text = extension["nosotros"]
    holders[17].text_frame.text = extension["vosotros"]
    holders[18].text_frame.text = extension["ellos/ellas/Ustedes"]

    holders[19].text_frame.text = EnglishVocabMeta.tense_info[tense]
    holders[20].text_frame.text = " ".join(["人称", person, "的变位"])

    note = slide.notes_slide
    note.notes_text_frame.text = v["word"]


  def create_multi_tiempo_verb_word(self, v):
    layout = self.prs.slide_layouts.get_by_name("Verb multi tiempo")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders
    
    variations = json.loads(v["variations"].replace("\'", "\""))

    origin, word, meaning = holders[10], holders[11], holders[12]
    origin.text_frame.text = variations["origin"]
    word.text_frame.text = v["word"]

    ms = v["meaning"].split(",")
    if len(ms) > 4:
      ms = ms[:4]

    meaning.text_frame.text = "\n".join(ms)
 
    formats = variations["formats"]
    holders[13].text_frame.text = "\n".join([EnglishVocabMeta.tense_info[f["tense"]] if "tense" in f.keys() else EnglishVocabMeta.tense_info[f["format"]] for f in formats])
    holders[14].text_frame.text = "\n".join([" ".join([f["person"], "的变位"]) if "person" in f.keys() else "" for f in formats])

    note = slide.notes_slide
    note.notes_text_frame.text = v["word"]

  def create_particle_verb_word(self, v):
    layout = self.prs.slide_layouts.get_by_name("Verb participle")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders
    
    variations = json.loads(v["variations"].replace("\'", "\""))

    origin, word, meaning = holders[10], holders[11], holders[12]
    origin.text_frame.text = variations["origin"]
    word.text_frame.text = v["word"]

    ms = v["meaning"].split(",")
    if len(ms) > 4:
      ms = ms[:4]

    meaning.text_frame.text = "\n".join(ms)
 
    sign = variations["formats"][0]

    holders[13].text_frame.text = EnglishVocabMeta.tense_info[sign["format"]]

    note = slide.notes_slide
    note.notes_text_frame.text = v["word"]

  def create_verb_word(self, v):
    layout = self.prs.slide_layouts.get_by_name("Original verb vocab")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders
    
    word, meaning = holders[11], holders[12]
    word.text_frame.text = v["word"]

    ms = v["meaning"].split(",")
    if len(ms) > 4:
      ms = ms[:4]

    meaning.text_frame.text = "\n".join(ms)
 
    if v["examples"] != "":
      examples = json.loads(v["examples"])
      if len(examples) == 2:
        original, translated = holders[14], holders[15]
        ex = examples[0]
        original.text_frame.text = ex["original"]
        translated.text_frame.text = ex["translated"]

        original, translated = holders[16], holders[17]
        ex = examples[1]
        original.text_frame.text = ex["original"]
        translated.text_frame.text = ex["translated"]

 
    note = slide.notes_slide
    note.notes_text_frame.text = v["word"]

  def create_vocab_group(self, vocabs):
    for v in vocabs:
      pos = v["dict_pos"]
      if pos in EnglishVocabMeta.pos_info['noun'][2] and v['extension'] != '' and v['examples'] != '[]':
        self.create_noun_word(v) 
      elif pos in EnglishVocabMeta.pos_info['adj'][2] and v['extension'] != '' and v['examples'] != '[]':
        self.create_adj_word(v)
      elif pos in EnglishVocabMeta.pos_info['verb'][2] and v['examples'] != '[]':
        self.create_verb_word(v)
      elif v['examples'] != '[]':
        self.create_default_word(v)
      else:
        pass
        #self.create_default_word(v)

  def create_vocab(self):
    """
    Create 3 pos group pages, which are noun, adj, verb, refer to EnglishVocabMeta
    """
    for index, (pos, ws) in enumerate(self.displayed_word_distribution.items()):
      if index < 3 and len(ws) > 0:
        subtitle = EnglishVocabMeta.pos_info[pos][0]
        title = EnglishVocabMeta.pos_info[pos][1].upper()
        self.create_vocab_title(index+1, title, subtitle)
        self.create_vocab_group(ws)

  def create_ending(self):
    layout = self.prs.slide_layouts.get_by_name("Thanks")
    self.prs.slides.add_slide(layout)


  def save_ppt(self, destfile):
    self.prs.save(destfile)

  def convert_to_ppt(self, destfile='test.pptx'):
    self.create_home()
    self.create_statistic()
    self.create_vocab()
    self.create_ending()

    self.save_ppt(destfile)    
    
