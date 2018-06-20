"Simple call to check real results."

import sys
sys.path.append('../presentation_text_replacer')

from presentation_text_replacer import pptRep

pptRep(
    inFile='../test_data/mel-in.pptx',
    outfile='../test_data/mel-out.pptx',
    repFile='../test_data/rep.ini'
)