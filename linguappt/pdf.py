from pdf2image import convert_from_path
from subprocess import call, DEVNULL
import os

class Pdf:
  """
  Generate pdf file from pptx file
  """

  def pptx2pdf(self, pptx, dstfolder='./'):
    self.dstfolder = dstfolder
    if not os.path.isdir(dstfolder):
      os.mkdir(self.dstfolder) 
    call(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', self.dstfolder, pptx], stdout=DEVNULL)
    self.pdfpath = os.path.join(dstfolder, os.path.basename(pptx).split('.')[0] + '.pdf')

  def __init__(self, pptx, dstfolder='./'):
    self.pptx2pdf(pptx, dstfolder)

  def save_as_images(self, start=0, end=None, imgfolder='./'):
    if not os.path.isdir(imgfolder):
      os.mkdir(imgfolder) 
    self.images = convert_from_path(self.pdfpath, thread_count=2, use_pdftocairo=True, size=(800, None), timeout=240)
    for index, image in enumerate(self.images[start:end]):
      image.save(os.path.join(imgfolder, str(index)+".jpg"))
    return len(self.images[start:end])


