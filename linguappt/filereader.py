import csv

def readCSV(filename):
  with open(filename) as csvfile:
    readCSV = csv.reader(csvfile, skipinitialspace=True, delimiter='\t')
    header = next(readCSV)
    return [dict(zip(header, row)) for row in readCSV]

