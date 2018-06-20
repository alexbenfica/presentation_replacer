"Simple call to check real results."

import sys
sys.path.append('../presentation_text_replacer')

from presentation_text_replacer import PptRep

ppt_rep = PptRep(
    input_file='../test_data/mel-in.pptx',
    output_file='../test_data/mel-out.pptx',
    replacer_file='../test_data/rep.ini'
)


ppt_rep.process_file()

