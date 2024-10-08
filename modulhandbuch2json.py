# Tool for extracting predefined text data from PDF documents
# specifically from "Modulhandbücher der Hochschule Bochum"
# by Christian Figge - Germany 2024
# (v04 - Work in progress)
#
# Only works on PDFs with plain trailer sections and xref tables
#
# Provided as-is. No warranties.
# Shared under CC0 (https://creativecommons.org/public-domain/cc0/)

import sys
import re
import zlib
import copy
import json

# an array of dicts will hold modul info:
modul_dummy = {
    "name": "DUMMY", # 0
    "modulnr": "", # 1
    "workload": 0, # 2
    "credits": 0, # 3
    "semester_start": 0, # 4
    "haeufigkeit": "", # 5 
    "dauer": 0, # 6
    "veranstaltungen": "", # 7
    "kontaktzeit": 0, # 8
    "selbststudium": 0, # 9
    "gruppen": "", # 10
    "lernergebnisse": "", # 11
    "inhalte": "", #12
    "lehrformen": "", # 13
    "teilnahmevoraussetzungen": "", # 14
    "pruefungsformen": "", # 15
    "wincondition": "", # 16
    "verwendung": "", # 17
    "stellenwert": "", # 18
    "modulbeauftragte": "", # 19
    "sonstiges": "", # 20
}

# words to scan the pages for (can be overwritten by providing a keywords.txt file)
# default keywords for Informatik B.Sc.: (might throw a SyntaxWarning you can safely ignore)
keywords = [
    "Ken.*?nummer", # 0
    "Workload", # 1
    "Credits", # 2
    "Studiensem", #3
    "Häufigkeit des Angebots", #4
    "Dauer", #5
    "Lehrveranstaltungen", # 6
    "Kontaktzeit.*?SWS", # 7
    "Selbststudium", # 8
    "geplante Gruppengröße", # 9
    "Lernergebnisse \(learning.*?outcomes\).*?Kompetenzen", # 10 
    "Inhalte", # 11
    "Lehrformen", # 12
    "Teilnahmevoraussetzungen", # 13
    "Prüfungsformen", # 14
    "Voraussetzungen für die Vergabe von Kreditpunkten", # 15
    "Verwendung des Moduls \(in anderen Studiengängen\)", # 16
    "Stellenwert der Note für die Endnote", # 17
    "Modulbeauftragte/r und hauptamtlich Lehrende", # 18
    "Sonstige Informationen" # 19
]
    
def getSearchPatterns():
    # get keywords from file
    if(len(sys.argv) > 3):
        try:
            with open(sys.argv[3], "rb") as kwFile:
                lines = kwFile.readlines()
                if(len(lines) != len(keywords)):
                    print(sys.argv[3] + " doesn't contain " + str(len(keywords)) + " keywords (1 per line)!\nAborting ...")
                    exit()
                for i in range(len(lines)):
                    keywords[i] = lines[i].decode("UTF-8").strip()
                
        except FileNotFoundError:
            print("File " + sys.argv[3] + " not found!")
            exit()
        
    print("Compiling regex patterns ... ", end="")
    patterns = [
        re.compile("(.*?) " + keywords[1]), # name
        re.compile(keywords[0] + "(.*?) " + keywords[1]), # kennnr
        re.compile(keywords[1] + ".*?([0-9]*?)[^0-9]*?" + keywords[2]), # workload
        re.compile(keywords[2] + ".*?([0-9]*?)[^0-9]*?" + keywords[3]), # credits
        re.compile(keywords[3] + ".*?([0-9]).*?" + keywords[4]), # semester start
        re.compile(keywords[4] + "(.*?) " + keywords[5]), # haeufigkeit
        re.compile(keywords[5] + ".*?([0-9]?) Sem"), # dauer
        re.compile(keywords[6] + "(.*?)" + keywords[7]), # veranstaltungen
        re.compile(keywords[7] + ".*?([0-9]*?)[^0-9]*?" + keywords[8]), # kontaktzeit
        re.compile(keywords[8] + ".*?([0-9]*?)[^0-9]*?" + keywords[9]), # selbststudium
        re.compile(keywords[9] + "(.*?)2 " + keywords[10]), # gruppengröße
        re.compile("2 " + keywords[10] + "(.*?)3 " + keywords[11]), # lernziele
        re.compile("3 " + keywords[11] + "(.*?)4 " + keywords[12]), # inhalte
        re.compile("4 " + keywords[12] + "(.*?)5 " + keywords[13]), # lehrformen
        re.compile("5 " + keywords[13] + "(.*?)6 " + keywords[14]), # teilnahmevoraus.
        re.compile("6 " + keywords[14] + "(.*?)7 " + keywords[15]), # prüfungsform
        re.compile("7 " + keywords[15] + "(.*?)8 " + keywords[16]), # wincondition
        re.compile("8 " + keywords[16] + "(.*?)9 " + keywords[17]), # verwendung
        re.compile("9 " + keywords[17] + "(.*?)10 " + keywords[18]), # stellenwert
        re.compile("10 " + keywords[18] + "(.*)"), # lehrende 
        re.compile("11 " + keywords[19] + "(.*)") # sonstiges   
    ]
    print("Done")
    return patterns
    

def pages2dict(pages):
    # TODO get info on first page: fachbereich, studiengang, PO, stand
    
    searchPatterns = getSearchPatterns()   
    modules = []
    pIdx = 0
    while (pIdx < len(pages)): 
        modul = copy.deepcopy(modul_dummy)
        pList = pages[pIdx]
        pText = "".join(pList)
        
        #print(pText) # debug
        
        # check if it's a module page
        if(keywords[1] in pText and keywords[2] in pText): # they got typos in keywords[0]...
            startPage = pIdx 
            print("Processing entry on p. " + str(startPage + 1) + ": ", end="")
            
            # include next page if module is a 2-pager
            while(keywords[-2] not in pText): 
                pIdx += 1
                pList += pages[pIdx]
                pText += "".join(pages[pIdx])
            
            # get module name
            nameIdx = getModulNameIdx(pList)
            if(nameIdx < 0):
                match = searchPatterns[0].search(pText)
                if(match):
                    name = getModulName(match.group(1))
                    if(len(name) < len(match.group(1))):
                        modul["name"] = name.strip(' ')
                    else:
                        modul["name"] = pList[1].strip(' ')
                        printUnsureWarning(startPage, "name", modul["name"])
            else:
                modul["name"] = pList[nameIdx].strip(' ')

            print('"' + modul["name"] + '" ... ')
            
            # get dict values for keys 1 - 18
            keys = list(modul.keys())
            lastPos = 0
            for i in range(1, 19):
                match = searchPatterns[i].search(pText, lastPos)
                if(match):
                    matchStr = match.group(1).lstrip('.:').strip(' ')
                   
                    if(isinstance(modul[keys[i]], int)):
                        if(len(matchStr) > 0):
                            modul[keys[i]] = int(matchStr)
                        else:
                            printUnsureWarning(startPage, keys[i], modul[keys[i]])
                    else:
                        modul[keys[i]] = matchStr
                    lastPos = match.end(1)
                else:
                    printUnsureWarning(startPage, keys[i], modul[keys[i]])
            
            # get beauftragte & sonstiges (falls vorhanden)
            match = searchPatterns[19].search(pText, lastPos)
            if(match):
                matchStr = match.group(1)
                endIdx = matchStr.find("11 " + keywords[19])
                if(endIdx != -1):
                    modul[keys[19]] = matchStr[:endIdx].lstrip('.:').strip(' ')
                    
                    # get sonstiges
                    match = searchPatterns[20].search(matchStr, endIdx)
                    modul[keys[20]] = match.group(1).lstrip('.:').strip(' ')
                else:
                    modul[keys[19]] = matchStr.lstrip('.:').strip(' ')
            else:
                printUnsureWarning(startPage, keys[19], modul[keys[19]])
                
            modules.append(modul)
            
        pIdx += 1
        
    dictModulhandbuch = {
        "fachbereich" : "",
        "studiengang" : "",
        "PO" : "",
        "stand" : "",
        "module" : modules
    }
        
    return dictModulhandbuch

def invalidPDF(fileIdx = 1, errorMsg = "Unknown error"):
    print("Error: Input file '" + sys.argv[fileIdx] + "' is incompatible! (" + errorMsg + ")\nAborting...")
    exit();
    
def printUnsureWarning(pageIdx, key, value):
    print(">> Please check keyword '" + key + "' : '" + str(value) + "'")
    
def getObjBytes(pdfFile, objOffsets, objIdx):
    pdfFile.seek(objOffsets[objIdx])
    objBytes = pdfFile.read()
    return objBytes[:objBytes.find(b'endobj')]

def getPlainText(byteString):
    return b''.join(re.findall(rb'\((.*?)\)', byteString))

def getGlyphText(byteString):
    # 16 bit glyph value (as 4 hex digits) + 29 = equivalent ascii char (mostly)
    hexString = b''.join(re.findall(rb'<([0-9A-Fa-f]*?)>', byteString))
    charList = [chr(29 + int(hexString[i:i+4], 16)) for i in range(0, len(hexString), 4)]

    # TODO proper en/decoding
    # even chardet couldn't detect the correct encoding,
    # so for now we just change the most critical symbols manually:
    for i in range(len(charList)):
        match charList[i]:
            case 'Ó':
                charList[i] = 'ö' 
            case 'Ù':
                charList[i] = 'ü'
            case '÷':
                charList[i] = '"'
            case 'õ':
                charList[i] = '"'
            case 'á':
                charList[i] = '"'
            case 'Ð':
                charList[i] = '"'
            case 'È':
                charList[i] = '...'
            case '\x89':
                charList[i] = 'ä'
            # TODO ß und capital umlaute
    
    return "".join(charList)
    
def getModulNameIdx(pList):
    for i in range(len(pList)):
        if(re.match(keywords[0], pList[i])):
            return i - 1
    return -1

def getModulName(matchStr):
    match = re.search("[0-9.]*? (.*?) ", matchStr)
    if(match):
        firstWord = match.group(1)
        
        # find idx of second instance and slice
        idx = matchStr.find(firstWord[:-1], match.end(1))
        if (idx != -1):
            return matchStr[match.start(1):idx]
 
    return matchStr

def main():
    if(len(sys.argv) >= 3):
        #TODO iterate over multiple input files
        
        encoding = 'cp1252' # TODO
        fileIdx = 1
        try:
            with open(sys.argv[fileIdx], "rb") as f:
                print("Fetching text contents ... ", end="")
                
                fBytes = f.read()
                fSize = f.tell()
                
                # get trailer info
                trailerPos = fBytes.rfind(b'trailer')
                
                if(trailerPos < 0):
                    invalidPDF(fileIdx, "trailer section not found")
                    
                f.seek(trailerPos)
                trailerBytes = f.read()
                nObjs = int(re.search(rb'/Size ([0-9]*?)[^0-9]', trailerBytes).group(1))
                rootObjIdx = int(re.search(rb'/Root ([0-9]*?)[^0-9]', trailerBytes).group(1))
                
                # get info about xref table(s)
                pos = fSize
                xrefInfo = []   # [firstObjIdx, numOfObjs, tablePos] for each table
                while ((pos := fBytes.rfind(b'startxref', 0, pos)) != -1): 
                    f.seek(pos) # jump to xref ref
                    next(f)     # skip 'startxref' tag
                    f.seek(int(f.readline().strip(b'\r\n'))) # jump to xref table
                    next(f)     # skip 'xref' tag
                    xrefInfo = [[int(s) for s in f.readline().strip(b'\r\n').split(b' ')] + [f.tell()]] + xrefInfo # prepend info
                    
                if (len(xrefInfo) == 0):
                    invalidPDF(fileIdx, "xref table not found")
                
                # get object offsets
                objOffsets = [-1 for _ in range(nObjs)]
                for xi in xrefInfo:
                    f.seek(xi[2])
                    objIdx = xi[0]
                    for o in range(xi[1]):
                        objOffsets[objIdx] = int(f.readline()[:10]) 
                        objIdx += 1
                
                # get page obj indices
                rootObjBytes = getObjBytes(f, objOffsets, rootObjIdx)
                pagesObjIdx = int(re.search(rb'/Pages ([0-9]*?)[^0-9]', rootObjBytes).group(1))
                pagesObjBytes = getObjBytes(f, objOffsets, pagesObjIdx)
                kidsBytes = re.search(rb'/Kids.*?\[([0-9 R]*?)\]', pagesObjBytes).group(1).strip().split(b' ')
                pageObjIndices = [int(kidsBytes[i]) for i in range(0, len(kidsBytes), 3)]
                          
                # get content obj indices
                contentObjIndices = []
                for pIdx in pageObjIndices:
                    pageBytes = getObjBytes(f, objOffsets, pIdx)
                    contentObjIndices.append( int(re.search(rb'/Contents ([0-9]*?)[^0-9]', pageBytes).group(1)) )
                
                # get page texts 
                pages = []
                for cIdx in contentObjIndices:
                    contentBytes = getObjBytes(f, objOffsets, cIdx)
                    if (re.search(rb'(?s)<<.*?/Filter.*?/FlateDecode.*?>>', contentBytes)): # needs decoding
                        streamBytes = re.search(rb'(?s)stream(.*?)endstream', contentBytes).group(1)[2:-2] #.strip(b'\r\n')
                        contentBytes = zlib.decompress(streamBytes)
                    else:
                        pass 
                    
                    contentBytes = contentBytes.replace(b'\\(', b'[').replace(b'\\)', b']') # substitute literal ()'s    
                    cells = re.findall(rb'(?s)<</MCID [0-9]*?.*?>> (.*?)\r\n EMC ', contentBytes) # get the page's table cells
                      
                    i = 0
                    while(i < len(cells)):
                        # relevant text data is in []
                        # with 16bit hexadecimal glyphs in <> that need to be decoded seperately
                        # and 8bit plain text in (). Single glyphs are ignored
                        
                        foundGlyphTxt = re.search(rb'<[0-9A-Fa-f]{8,}?>', cells[i]) 
                        foundPlainTxt= re.search(rb'\((.*?)\)', cells[i])
                            
                        if(foundGlyphTxt and foundPlainTxt):
                            textParts = re.findall(rb'\[(.*?)\]', cells[i])
                                
                            # each part in [] appears to contain either glyphs or plaintext, never both
                            for j in range(len(textParts)):
                                
                                if(re.search(rb'<[0-9A-Fa-f]*?>', textParts[j])):
                                    textParts[j] = getGlyphText(textParts[j])
                                else:
                                    textParts[j] = getPlainText(textParts[j]).decode(encoding)
                                   
                            cells[i] = ''.join(textParts)
                               
                        elif(foundGlyphTxt): # only glyphs
                            cells[i] = getGlyphText(cells[i])
                        elif(foundPlainTxt): # only plaintext
                            cells[i] = getPlainText(cells[i]).decode(encoding)                            
                        else:
                            cells[i] = '' # clear anything else
                            
                        if(len(cells[i]) == 0 or len(cells[i].strip(' ')) == 0):
                            del cells[i]
                        else:
                            cells[i] = cells[i].replace('[', '(').replace(']',')')
                            i += 1
                        
                    pages.append(cells)
                print("Done")
                
                dictModulhandbuch = pages2dict(pages)
                
                print("Writing output file '" + sys.argv[2] + "' ... ", end="")
                try:
                    with open(sys.argv[2], "w") as outFile:
                        json.dump(dictModulhandbuch, outFile)
                    print("Done")
                except FileNotFoundError:
                    print("\nERROR: File " + sys.argv[2] + " could not be accessed!\nExiting ...")
                    exit()
                
        except FileNotFoundError:
            print("\nERROR: File " + sys.argv[fileIdx] + " not found!\nExiting ...")
            exit()
        
    else:
        print(
            "\nMissing arguments!\n" + \
            "\n" + \
            "Usage: " + sys.argv[0] + " INPUTFILE.pdf OUTPUTFILE.json [keywords.txt]")
            
if __name__ == '__main__':
    main()
