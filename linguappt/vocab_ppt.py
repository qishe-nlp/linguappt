from pptx import Presentation
from linguappt.filereader import readCSV
import json

class VocabPPT:
  """DON'T use this class directly
  """
  def __init__(self, sourcefile, title="", genre="classic"):
    self.template = self.__class__.templates[genre]
    self.prs = Presentation(self.template)
    self.title = title
    self.sourcefile = sourcefile
    self._assert_content()
    self._divide_by_pos()

  def _divide_by_pos(self):
    """Store words in self.word_distribution as dict, whose key is in ``EnglisgVocabMeta.pos_info`` and value is a list of words
    """
    word_distribution  = {}

    for e in self.content:
      key = self.__class__.metainfo.get_pos(e["dict_pos"])
      if key in self.__class__.ALLOWED_POSES:
        if key not in word_distribution:
          word_distribution[key] = []
        word_distribution[key].append(e)
    self.word_distribution = word_distribution


  def _assert_content(self):
    """
      Ensure the content has keys defined in ppt class
    """
    content = readCSV(self.sourcefile) 
    for e in content:
      keys = e.keys()
      assert(len(keys) == len(self.__class__.content_keys))
      for k in keys:
        assert(k in self.__class__.content_keys)
    self.content = content

  def _create_opening(self):
    layout = self.prs.slide_layouts.get_by_name("Opening for chinese")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders
 
    title = holders[10] 
    title.text_frame.text = self.title
    subtitle = holders[11]
    subtitle.text_frame.text = '词汇总结'

  def _create_statistics(self):
    layout = self.prs.slide_layouts.get_by_name("Statistics")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders

    for index, (pos, ws) in enumerate(self.word_distribution.items()):
      pos_name = holders[11+2*index]
      num = holders[10+2*index]
      pos_name.text_frame.text = self.__class__.metainfo.get_pos_name(pos).upper() 
      num.text_frame.text = str(len(ws))

  def _create_ending(self):
    layout = self.prs.slide_layouts.get_by_name("Thanks")
    self.prs.slides.add_slide(layout)

  def _create_vocab_title(self, title_content, subtitle_content):
    layout = self.prs.slide_layouts.get_by_name('Title for pos')
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders

    title = holders[10]
    title.text_frame.text = title_content
    subtitle = holders[11]
    subtitle.text_frame.text = subtitle_content
    note = slide.notes_slide
    note.notes_text_frame.text = title_content.lower()

  def _create_default_word(self, v):
    layout = self.prs.slide_layouts.get_by_name("Default")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders

    pos, noun, meaning = holders[10], holders[11], holders[12]
    pos.text_frame.text = v["dict_pos"]
    noun.text_frame.text = v["word"]
    ms = v["meaning"].split(",")
    if len(ms) > 4:
      ms = ms[:4]
    meaning.text_frame.text = "\n".join(ms) 

    note = slide.notes_slide
    note.notes_text_frame.text = v["word"]

  def _create_default_word_with_examples(self, v):
    layout = self.prs.slide_layouts.get_by_name("Default with examples")
    slide = self.prs.slides.add_slide(layout)
    holders = slide.shapes.placeholders

    pos, noun, meaning = holders[10], holders[11], holders[12]
    pos.text_frame.text = v["dict_pos"]
    noun.text_frame.text = v["word"]
    ms = v["meaning"].split(",")
    if len(ms) > 4:
      ms = ms[:4]
    meaning.text_frame.text = "\n".join(ms) 

    examples = json.loads(v["examples"])
    if len(examples) >= 1:
      original, translated = holders[13], holders[14]
      ex = examples[0]
      original.text_frame.text = ex["original"]
      translated.text_frame.text = ex["translated"]
    if len(examples) >= 2:
      original, translated = holders[15], holders[16]
      ex = examples[1]
      original.text_frame.text = ex["original"]
      translated.text_frame.text = ex["translated"]

    note = slide.notes_slide
    note.notes_text_frame.text = v["word"]


  def _create_vocab_group(self, pos, vocabs):
    for v in vocabs:
      if pos == "noun":
        self._create_noun(v) 
      elif pos == "adj":
        self._create_adj(v)
      elif pos == "verb":
        self._create_verb(v)
      #else:
      #  pass


  def _create_vocab(self):
    """Create vocab group slides, which are noun, adj, verb, restrained by ALLOWED_POSES
    """
    for pos, ws in self.word_distribution.items():
      if len(ws) > 0:
        subtitle = self.__class__.metainfo.get_pos_cn_name(pos)
        title = self.__class__.metainfo.get_pos_cn_name(pos).upper()
        self._create_vocab_title(title, subtitle)
        self._create_vocab_group(pos, ws)

  def save_ppt(self, destfile):
    self.prs.save(destfile)

  def convert_to_ppt(self, destfile='test.pptx'):
    self._create_opening()
    self._create_statistics()
    self._create_vocab()
    self._create_ending()

    self.save_ppt(destfile)    

  def _create_noun(self, v):
    raise Exception("{} has to be implemented in subclass".format("_create_noun"))

  def _create_verb(self, v):
    raise Exception("{} has to be implemented in subclass".format("_create_verb"))

  def _create_adj(self, v):
    raise Exception("{} has to be implemented in subclass".format("_create_adj"))


