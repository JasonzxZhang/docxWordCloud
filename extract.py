import os
import glob
from docx import Document
from wordcloud import WordCloud, STOPWORDS
from gensim.summarization import summarize

def main():
	filePath = getDir()
	answerList = ""
	for file in os.listdir(filePath):
		if file.endswith(".docx"):
			answerList += getContent(os.path.join(filePath, file))
			answerList += "\n"
	writeFile(answerList, filePath)
	createSummary(answerList)
	createWordCloud(answerList, filePath)
			
def createWordCloud(text, filePath):
	"""Creates a WordCloud with input text by extracting most frequent
	words from the input text. Then writes the generate image to input
	filepath directory.

	Parameters
	----------
	text : string
		Raw string containing all the textual content
	filePath : string
		Description of arg2

	Returns
	-------
	None
		No objects are returned
	"""
	# STOPWORDS include nonmeaningful words such as 'I', 'is', etc.
	stp = set(STOPWORDS)

	# Creating a platte of wordcloud canvas, you may change 
	# different parameters as needed
	wc = WordCloud(	background_color = "white",
					max_words = 100,
					stopwords = stp,
					collocations=False,
					width=800,
					height=400)
	wc.generate(text = text)
	wc.to_file(os.path.join(filePath, "worldcloud.png"))

def createSummary(text):
	# TODO: an additional feature to work on using Machine Learning
	sumText = summarize(text=text)
	print(sumText)

def writeFile(text, filePath):
	# TODO: helper function for summary creation
	text_file = open(os.path.join(filePath, "summary.txt"), "w")
	text_file.write(text)
	text_file.close()

def getContent(filename=None):
	"""Extract text from a '.docx' file from a selected range -
	'startIndex' to 'delim'.

	Parameters
	----------
	filename : string
		path to the docx file

	Returns
	-------
	string
		text content extract from the range
	"""
	if not filename:
		print("File field may not be left blank!")
		return

	doc = Document(filename)
	startIndex = getSearchIndex()
	delim = getDeliminator()
	# You may also manually hard code search start and end index
	# startIndex = "xxxxxx"
	# delim = "xxxxxx"
	isTarget = False
	text = ""
	for paragraph in doc.paragraphs:
		if startIndex in paragraph.text:
			isTarget = True
			pass
		if delim in paragraph.text:
			isTarget = False
		if isTarget:
			text += paragraph.text
	return text
		

def getDir():
	"""Get root file directory from user input.

	Parameters
	----------
	None

	Returns
	-------
	string
		Root file directory path
	"""

	rootDir = input("Input the path for docx text extraction -->")
	# TOOD: implement loop, remove trailing space
	if os.path.exists(rootDir):
		return rootDir
	else:
		print("INVALID PATH!")

def getSearchIndex():
	"""Get beginning of the search content from user input.

	Parameters
	----------
	None

	Returns
	-------
	string
		Search index
	"""

	search = input("Input the question or start of the text -->")
	# TOOD: implement loop, remove trailing space
	if search == "":
		print("INVALID SEARCH INPUT!")
	else:
		return search

def getDeliminator():
	"""Get end of the search content from user input.

	Parameters
	----------
	None

	Returns
	-------
	string
		Deliminator of the desired text
	"""

	delim = input("Input the path for docx text extraction -->")
	# TOOD: implement loop, remove trailing space
	if delim == "":
		print("INVALID DELIMINATOR INPUT!")
	else:
		return delim


if __name__ == "__main__":
	main()