from pdf2image import convert_from_path
from subprocess import call, DEVNULL
import os

def pptx2pdf(pptx, pdffolder='./'):
  if not os.path.isdir(pdffolder):
    os.mkdir(pdffolder) 
  call(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', pdffolder, pptx], stdout=DEVNULL)
  pdfpath = os.path.join(pdffolder, os.path.basename(pptx).split('.')[0] + '.pdf')
  return pdfpath

def pdf2images(pdfpath, imgfolder='./', start=0, end=None):
  if not os.path.isdir(imgfolder):
    os.mkdir(imgfolder) 
  images = convert_from_path(pdfpath, thread_count=2, use_pdftocairo=True, size=(800, None), timeout=240)
  for index, image in enumerate(images[start:end]):
    image.save(os.path.join(imgfolder, str(index)+".jpg"))
  return len(images[start:end])


