from linguappt import SpanishVocabPPT, Pdf
from linguappt import __version__
import click
import json
import os
 
def print_version(ctx, param, value):
  if not value or ctx.resilient_parsing:
      return
  click.echo(__version__)
  ctx.exit()
 
@click.command()
@click.option("--sourcecsv", prompt="source csv file path", help="Specify the source csv file path")
@click.option("--title", prompt="title of the pptx", help="Specify the title of the pptx")
@click.option("--destpptx", default="test.pptx", prompt="destination pptx file", help="Specify the destination pptx file name")
def vocab_csv2ppt(sourcecsv, title, destpptx):
  phase = {"step": 1, "msg": "Start ppt generation"}
  print(json.dumps(phase))

  vp = SpanishVocabPPT(sourcecsv, title)
  vp.convert_to_ppt(destpptx)

  phase = {"step": 2, "msg": "Finish ppt generation"}
  print(json.dumps(phase))


@click.command()
@click.option("--sourcepptx", prompt="source pptx file path", help="Sepcify the source pptx file path")
@click.option("--destdir", prompt="dest pdf and pictures directory", help="Sepcify the pdf and picture destionation directory")
def ppt2pdf(sourcepptx, destdir):
  phase = {"step": 1, "msg": "Start pdf generation"}
  print(json.dumps(phase))

  pdf = Pdf(sourcepptx, destdir)

  phase = {"step": 2, "msg": "Finish pdf generation"}
  print(json.dumps(phase))

  phase = {"step": 3, "msg": "Start images generation"}
  print(json.dumps(phase))

  images_len = pdf.save_as_images(4, 4+6, destdir)
   
  phase = {"step": 4, "msg": "Finish images generation", "images_len": images_len}
  print(json.dumps(phase))


@click.command()
@click.option('--version', is_flag=True, callback=print_version, expose_value=False, is_eager=True)
@click.option("--sourcecsv", prompt="source csv file path", help="Specify the source csv file path")
@click.option("--name", default="test", prompt="output file name", help="Specify the file name")
@click.option("--pptxdir", prompt="dest pptx directory", help="Specify the pptx destination directory")
@click.option("--pdfdir", prompt="dest pdf directory", help="Sepcify the pdf destionation directory")
@click.option("--imgdir", prompt="dest image directory", help="Sepcify the preview image destionation directory")
def vocab_csv2pptpdf(sourcecsv, name, pptxdir, pdfdir, imgdir):
  phase = {"step": 1, "msg": "Start ppt generation"}
  print(json.dumps(phase), flush=True)

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

  phase = {"step": 2, "msg": "Finish ppt generation, start pdf generation"}
  print(json.dumps(phase), flush=True)

  if not os.path.isdir(pdfdir):
    os.mkdir(pdfdir) 

  pdf_pro_dir = pdfdir + '/pro/'
  pdf_water_dir = pdfdir + '/water/'
  pdf = Pdf(pptx, pdf_pro_dir)
  watermark_pdf = Pdf(watermark_pptx, pdf_water_dir)

  phase = {"step": 3, "msg": "Finish pdf generation, start images generation"}
  print(json.dumps(phase), flush=True)

  images_len = pdf.save_as_images(0, 6, imgdir)
   
  phase = {"step": 4, "msg": "Finish images generation", "images_len": images_len}
  print(json.dumps(phase), flush=True)

