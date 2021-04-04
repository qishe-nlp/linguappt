from pptx import Presentation
from linguappt.filereader import readCSV
import json

class PhrasePPT:
  """DON'T use this class directly
  """
  def __init__(self, sourcefile, title="", genre="classic"):
    self.template = self.__class__.templates[genre]
    self.prs = Presentation(self.template)
    self.title = title
    self.sourcefile = sourcefile
    self._assert_content()

  def _assert_content(self):
    """Ensure the content has keys defined in ppt class
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
    subtitle.text_frame.text = '短语总结'

  def _create_ending(self):
    layout = self.prs.slide_layouts.get_by_name("Thanks")
    self.prs.slides.add_slide(layout)

  def save_ppt(self, destfile):
    self.prs.save(destfile)

  def convert_to_ppt(self, destfile='test.pptx'):
    self._create_opening()
    self._create_phrase()
    self._create_ending()

    self.save_ppt(destfile)    

