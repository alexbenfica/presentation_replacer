#!/usr/bin/python -tt
# -*- coding: utf-8 -*-
import os
import argparse

# This module will parse each argument

parser = argparse.ArgumentParser(
    description="Replace a list of text strings on presentations",
    epilog="Set input and ouput files and a txt file with strings to be replaced")

parser.add_argument("--inF" ,help="Input presentation file.",)
parser.add_argument("--outF",help="Output presentation file.",)    
parser.add_argument("--repF",help="Txt (ini format) file with strings to be replaced",)    

args = parser.parse_args()