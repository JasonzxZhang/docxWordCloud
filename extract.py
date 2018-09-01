import os
import glob
from docx import Document
from wordcloud import WordCloud, STOPWORDS
from gensim.summarization import summarize

def main():
	# filePath = getDir()
	filePath = "/Users/jason/Documents/Syracuse/ECS102_Grading/Section2_lab1"
	# answerList = list()
	answerList = ""
	for file in os.listdir(filePath):
	    if file.endswith(".docx"):
	        # answerList.append(getAnswer(os.path.join(filePath, file)))
	        answerList += getAnswer(os.path.join(filePath, file))
	        answerList += "\n"
	writeFile(answerList, filePath)
	createSummary(answerList)
	createWordCloud(answerList, filePath)
	        
def createWordCloud(text, filePath):
	stp = set(STOPWORDS)
	wc = WordCloud(	background_color = "white",
					max_words = 100,
					stopwords = stp,
					collocations=False,
					width=800,
					height=400)
	wc.generate(text = text)
	wc.to_file(os.path.join(filePath, "worldcloud.png"))

def createSummary(text):
	sumText = summarize(text=text)
	print(sumText)

def writeFile(text, filePath):
	text_file = open(os.path.join(filePath, "summary.txt"), "w")
	text_file.write(text)
	text_file.close()

def getAnswer(filename=None):
	if not filename:
		print("File field may not be left blank!")
		return

	doc = Document(filename)
	# startIndex = getSearchIndex()
	# delim = getDeliminator()
	startIndex = "Please write and submit a paragraph describing what you expect to learn in ECS 102:"
	delim = "ECS 102	Fall 2018"
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
	rootDir = input("Input the path for docx text extraction -->")
	# TOOD: implement loop, remove trailing space
	if os.path.exists(rootDir):
		return rootDir
	else:
		print("INVALID PATH!")

def getSearchIndex():
	search = input("Input the question or start of the text -->")
	# TOOD: implement loop, remove trailing space
	if search == "":
		print("INVALID SEARCH INPUT!")
	else:
		return search

def getDeliminator():
	delim = input("Input the path for docx text extraction -->")
	# TOOD: implement loop, remove trailing space
	if delim == "":
		print("INVALID DELIMINATOR INPUT!")
	else:
		return delim


if __name__ == "__main__":
	main()