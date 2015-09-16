try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile

"""
Module that converts text from MS XML Word document (.docx) to CSS styled HTML.
From
"""

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
RUNSTYLE = WORD_NAMESPACE + 'r'
RUNPROPS = WORD_NAMESPACE + 'rPr'
PARAPROPS = WORD_NAMESPACE + 'pPr'
PARASTYLE = WORD_NAMESPACE + 'pStyle'
fontSize = WORD_NAMESPACE + 'sz'
justification = WORD_NAMESPACE + 'jc'
indent = WORD_NAMESPACE + 'ind'
shdBreak = WORD_NAMESPACE + 'shd'
shade = WORD_NAMESPACE + 'shd'
paraBorders = WORD_NAMESPACE + 'pBdr'
borderTypes = {"single":"solid","dashed":"dashed","dotted":"dotted","dashSmallGap":"dashed","double":"double"}
def get_docx_text(path):
    #use white-space: pre CSS
    """
    Take the path of a docx file as argument, return the formatted text.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)
    
    #elements = list()
    styleAdditions = list()
    elementString = ["<article>"]
	
    for paragraph in tree.iter(PARA):
        """for paraStyleElement in paragraph.find(PARAPROPS).iter():
            if paraStyleElement =="""
		
		#JUSTIFICATION
        justificationInfo = paragraph.find(PARAPROPS).find(justification).attrib[WORD_NAMESPACE+"val"]
        if justificationInfo != "left" and justificationInfo != "both":
            styleAdditions.append("text-align:"+justificationInfo+";")
        elif justificationInfo == "both":
            styleAdditions.append("text-align:justify;")
		
		#INDENT
        if paragraph.find(PARAPROPS).find(indent):
            indentInfo = dict()
            for attribKey in paragraph.find(PARAPROPS).find(indent).attrib:
                if attribKey == WORD_NAMESPACE +"hanging":
                    indentInfo["hanging"] = (float(paragraph.find(PARAPROPS).find(indent).attrib[attribKey])/20)
                if attribKey == WORD_NAMESPACE +"end" or attribKey == WORD_NAMESPACE +"right":
                    indentInfo["right"] = (float(paragraph.find(PARAPROPS).find(indent).attrib[attribKey])/20)
                if attribKey == WORD_NAMESPACE +"start" or attribKey == WORD_NAMESPACE +"left":
                   indentInfo["left"] = (float(paragraph.find(PARAPROPS).find(indent).attrib[attribKey])/20)
            for key in indentInfo:
                if key == "hanging":
                    styleAdditions.append("margin-top:"+str(indentInfo[key])+";")
                if key == "right" or key == "end":
                    styleAdditions.append("margin-right:"+str(indentInfo[key])+";")
                if key == "left" or key == "start":
                    styleAdditions.append("margin-left:"+str(indentInfo[key])+";")
                #left => start, right => end, hanging
				
		#BACKGROUND COLOURING
        if paragraph.find(PARAPROPS).find(shade):
                if paragraph.find(PARAPROPS).find(shade).attrib[WORD_NAMESPACE +"fill"] != "auto":
                    styleAdditions.append("background-color: #"+paragraph.find(PARAPROPS).find(shade).attrib["fill"]+";")
                
		#BORDERS
        if paragraph.find(PARAPROPS).find(paraBorders):
            for sideElement in paragraph.find(PARAPROPS).find(paraBorders).iter():
                pass
                #if sideElement != WORD_NAMESPACE + "between":
                    #sideElement.attrib["val"] =
                        
		#GET ALL THE TEXT, FIX TO PUT TEXT INTO ITS HTML ELEMENT
        paragraphText = [element.text for element in paragraph.iter(TEXT) if element.text] #make a list comprehension of styles to match
        if not paragraphText:
            if paragraph.iter(shdBreak):            
                for shadow in paragraph.iter(shdBreak): #bloody generator making me use a for loop for one element
                    if shadow.attrib[WORD_NAMESPACE+"fill"] == "auto":
                        for element in paragraph.iter(fontSize):
                            breakSize = str(float(element.attrib[WORD_NAMESPACE+"val"])/2) #gets a dict of fontSize element attributes, then gets the value representing font size, in points
                            break
                        styleAdditions.append("min-height:"+breakSize+"pt;") #why not em? ATTENTION ATTENTION ATTENTION ATTENTION ATTENTION
                        styleAdditions.append("margin:0;")
                    break
        else:
            elementString.append("<p")
            elementString.append(' style="')
            #elementString.append("white-space:pre-wrap;word-wrap:break-word;")
            for style in styleAdditions:
                elementString.append(style)
            if ''.join(elementString)[-1:] != ";":
                elementString.append(";")
            
            elementString.append('"')
            elementString.append(">")
            elementString.append("</p>\n")
            styleAdditions = list()
        """texts = [node.text
                 for node in paragraph.iter(TEXT)
                 if node.text]
        if texts:
            elements.append(''.join(texts))"""
    elementString.append("</article>")
    return ''.join(elementString)
print get_docx_text("test2.docx")