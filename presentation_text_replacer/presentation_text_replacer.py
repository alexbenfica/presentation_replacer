"""
Replace a list of text strings on presentations
"""
import os
import sys
import pprint as pp
import configparser

from pptx import Presentation

class pptRep():
    def __init__(self, inFile, outfile, repFile):
        """
        :param inFile: Input presentation file
        :param outfile: Output presentation file.
        :param repFile: Txt (ini format) file with strings to be replaced
        """
        inFile = os.path.abspath(inFile)                
        if not os.path.isfile(inFile):
            exit('Input file does not exists: %s' % inFile)        
        self.prs = Presentation(inFile)        
        self.loadReplaceFile(repFile)
        self.doReplaces()
        self.savePres(outfile)

    def loadReplaceFile(self, repFile):
        repFile = os.path.abspath(repFile)                
        if not os.path.isfile(repFile):        
            exit('File with string to replace does not exists: %s' % repFile)
        self.cp = configparser.ConfigParser()
        self.cp.read(repFile)                    

    # retrieves a configuration value from config. file or the same text if it is not found
    def getReplace(self, txtToReplace):
        # the section name does not matter. use the first
        if len(self.cp.sections()):
            section = self.cp.sections()[0]

        searchTxt = txtToReplace.replace('{','').replace('}','').strip()
        if self.cp.has_option(section, searchTxt):
            rep = self.cp.get(section, searchTxt)            
            return rep
        
        return txtToReplace
        
    def savePres(self, outputFile):
        outputFile = os.path.abspath(outputFile)                
        self.prs.save(outputFile)
        
    
    def doReplaces(self):        
        '''
        Replaces each occurrence of text inside presentation file.
        '''
        for slide in self.prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame: continue
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if '{' in run.text:
                            if '}' in run.text:
                                pp.pprint(run.text)
                                # replaces if fould, leave it if not found!
                                run.text = self.getReplace(run.text)
