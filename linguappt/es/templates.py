import os

class Template:
  template_dir = os.path.dirname(__file__)
  classic = os.path.join(template_dir, 'templates/vocab_spanish_classic.pptx')
  # light = os.path.join(template_dir, 'templates/vocab_spanish_light.pptx')
  # dark = os.path.join(template_dir, 'templates/vocab_spanish_dark.pptx')
  watermark = os.path.join(template_dir, 'templates/vocab_spanish_watermark.pptx')

