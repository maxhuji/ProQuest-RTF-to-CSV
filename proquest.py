#This python code converts an RTF document containing >1000 news articles downloaded via ProQuest into a standardized CSV file. For every article it returns a row in the excel file extracting all the relevant metadata and content from every article.
#To run the code, open the command prompt, navigate to the folder in which the python script and the ProQuest documents are stored. Then run the script via "Python proquest.py"

import codecs
import glob

# Get a list of all RTF files with filenames "ProQuest"
g=glob.glob('ProQuest*.rtf')

# Open the output CSV file in binary write mode with UTF-8 encoding
o=codecs.open('proquest.csv',mode='wb',encoding='utf-8')

# Iterate over each RTF file
for f in g:
	i=codecs.open(f,encoding='mbcs')
	t=i.read()
	i.close()

	while '\\plain\\f1\\fs44\\cf0' in t:
		t=t[t.find('\\plain\\f1\\fs44\\cf0')+19:]
		title=t[:t.find('\\f1\\fs44\\cf0')]
		title=title.replace('\\f0\\fs20\\cf0','')
		title=title.replace(' \\par ',' ||| ')

		author=t[t.find('\\plain\\f1\\fs22\\cf0')+19:]
		author=author[:author.find('\\f1\\fs22\\cf0')]

		body=t[t.find('FULL TEXT\\b0\\par\\pard\\plain\\s0\\fi0\\li0\\ri0\\sa200\\sl320\\plain\\f1\\fs20\\cf0 \\u160?\\f1\\fs20\\cf0 \\par \\f1\\fs20\\cf0'):]
		body=body[:body.find('DETAILS')]
		body=body.replace('\\f1\\fs20\\cf0','')
		body=body.replace(' \\par ',' ||| ')
		body=body.replace('FULL TEXT\\b0\\par\\pard\\plain\\s0\\fi0\\li0\\ri0\\sa200\\sl320\\plain \\u160?','')
		body=body.replace('\\par\\pard\\plain\\s0\\fi0\\li0\\ri0\\sb200\\sa300\\sl320\\plain\\f1\\fs26\\b\\cf0','')

		outlet=t[t.find('Publication title: \\b0\\cell\\pard\\plain\\intbl\\s0\\ql\\fi0\\li0\\ri0\\sb80\\sa100\\plain\\f1\\fs20\\cf0')+92:]
		outlet=outlet[:outlet.find(';')]
		if '\\f0\\fs\\cf0' in outlet:
			outlet=outlet[:outlet.find('\\f0\\fs20\\cf0')]

		date=t[t.find('Publication date: \\b0\\cell\\pard\\plain\\intbl\\s0\\ql\\fi0\\li0\\ri0\\sb80\\sa100\\plain\\f1\\fs20\\cf0')+91:]
		date=date[:date.find(';')]
		if '\\f0\\fs\\cf0' in date:
			date=date[:date.find('\\f0\\fs20\\cf0')]
		if '\\f1\\fs20\\cf0' in date:
			date=date[:date.find('\\f1\\fs20\\cf0')]
		date=date.replace('Jan','01').replace('Feb','02').replace('Mar','03').replace('Apr','04').replace('May','05').replace('Jun','06').replace('Jul','07').replace('Aug','08').replace('Sep','09').replace('Oct','10').replace('Nov','11').replace('Dec','12')
		date=date.strip()
		date=date[date.find(',')-2:date.find(',')].strip()+'-'+date[:2]+'-'+date[-4:]
		date=date[:-4]+'20'+date[-2:]
		if len(date.strip())==9:
			date='0'+date.strip()

		section=t[t.find('Section: \\b0\\cell\\pard\\plain\\intbl\\s0\\ql\\fi0\\li0\\ri0\\sb80\\sa100\\plain\\f1\\fs20\\cf0')+82:]
		section=section[:section.find(';')]
		if '\\f0\\fs\\cf0' in section:
			section=section[:section.find('\\f0\\fs20\\cf0')]
		if '\\f1\\fs20\\cf0' in section:
			section=section[:section.find('\\f1\\fs20\\cf0')]

		url=t[t.find('Document URL: \\b0\\cell\\pard\\plain\\intbl\\s0\\ql\\fi0\\li0\\ri0\\sb80\\sa100\\plain\\f1\\fs20\\cf0')+87:]
		url=url[:url.find(';')]
		if '\\f0\\fs\\cf0' in url:
			url=url[:url.find('\\f0\\fs20\\cf0')]
		if '\\f1\\fs20\\cf0' in url:
			url=url[:url.find('\\f1\\fs20\\cf0')]
# Write the extracted information to the output CSV file
		o.write(outlet+'\t'+date+'\t'+section+'\t'+url+'\t'+author+'\t'+title+'\t'+body+'\n')
		print '.',
		
o.close()
