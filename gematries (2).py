#!/usr/bin/env python
# geany_encoding=UTF-8 #
#import requests
import os
import string
from os import listdir
from os.path import basename, isdir, isfile, join
import re
import datetime
#import pytz#for timezone
import swisseph as swe
from odf import text, teletype
from odf.opendocument import load
import codecs
import math

import pyexcel 
from pyexcel_ods import get_data
from pyexcel_ods import save_data
from collections import OrderedDict
import datetime



output_file=u'bible_ia_kabbalistic_search_.xls'
data_book = {'Sheet1':[]}
##pyexcel.save_as(array=data, dest_file_name=output_file)
#sheet = pyexcel.Sheet(data)
other_book = pyexcel.Book(data_book)
other_book.save_as(output_file)

input_file=u"xls_wordlist.xls"
empty_row=["NA","NA","NA",
			"NA","NA","NA",
			"NA","NA","NA"
			]
i=-1
row_to_add=[]
for value in empty_row:
	i+=1
	row_to_add.append(empty_row[i])
	
ephe_path='/usr/share/libswe/:/usr/local/share/libswe/'
swe.set_ephe_path(ephe_path)
gregflag=1
geolon=0.34#bordeaux
geolat=44.50#bordeaux

#Ce serait bien de faire une liste du top 10 des occurences contenant le moins de lettres 
# implement limit to psalms

def Read_cell(x=1,y=1):#this procedure uses an xls file and pyexcel the other should be harmonized
	print "reading cell",x,y,"from",input_file
	global input_file
	
	sheet = get_data(input_file)#another way to load a sheet this time in an ordered dictionary
	value=sheet["Sheet1"][y-1][x-1]
	return value

def blind_add_row(row):
	
	global output_file
#	print "adding row"
	book = pyexcel.get_book(file_name=output_file)#loads a sheet in a sheet object that can be modified
	book.Sheet1.row+= row
	book.save_as(output_file)
	

def new_day_ow(row):
	
	global output_file
	add=False
	
	#sheet = get_data(output_file)#another way to load a sheet this time in an ordered dictionary
	
	#checking if rows have already been added
	date=datetime.datetime.strftime(datetime.datetime.now(),"%d/%m/%Y")
	sheet = get_data(output_file)#this is rather a book than a sheet
	list_date=get_string_coord_column(sheet["Sheet1"],0, date)#note that the searching procedure takes a sheet not a book
	print "today's date offsets:",list_date
	if list_date==[]:
		add=True
		print "no today's date occurence found, adding row"
	else:
		print list_date
		print "inserting cells"
		print "this is the long procedure, adding",row
		i=0
		occurences=list_date
		for cell in row:
			i+=1
			#print "inserting at x",occurences[0][0]+i
			#print "type",type(cell)
			y=int(occurences[len(occurences)-1][1])
			x=int(occurences[0][0])+i
			sheet["Sheet1"][y-1][x-1]=cell
			# the following could be used to convert str to unicode
			#if (type(cell)==unicode or type(cell)==str):
			#	print "inserting unicode"
			#	sheet[y-1][x-1]=unicode(cell)
				#Insert_cell(occurences[0][0]+i,occurences[len(occurences)-1][1],unicode(cell))#inserting at the last occurence of the date
			#if (type(cell)==int):
			#	print "inserting int"
			#	sheet[y-1][x-1]=int(cell)
				#Insert_cell(occurences[0][0]+i,occurences[len(occurences)-1][1],int(cell))
		pyexcel.save_book_as(bookdict=sheet,dest_file_name=output_file)#saves a sheet
		print "Book recorded"
	
	if add:
	
		book = pyexcel.get_book(file_name=output_file)#loads a sheet in a sheet object that can be modified
		book.Sheet1.row+= row
		print "saving xls"
		book.save_as(output_file)

def reduc_theo(integer,limit):#number to be reduced, how many times it should be re done
	divinenew=0
	divine=0
	passage=0
	#print "reduction theosophique (esprit)de ",integer,":"
	while len(str(integer))>1 and passage<limit:
		passage=passage+1
		numletterd=len(str(integer))
		for chiffre in range (numletterd):
			isol=int(str(integer)[chiffre])
			if isol > 0 : 
				divinenew=divinenew+isol
		integer=divinenew
		#print"etape ",passage,":", divinenew
		divinenew=0
	#print divine
	return integer

def heure_dete(month, day):# lazy procedure for summer/winter time (it would be better to use pytz)
	if month>3 and month<10: return 1
	else: return 0
 
def daylight(Sun,Asc):#should return true when the Sun is above the horizon but would work if calculated in radial logic

	Des=Asc+180
	if Des>360: Des=Des-360
	if Des>180 : vernalpointdaylight=True
	else:
		vernalpointdaylight=False
	attrib=False
	if vernalpointdaylight and Asc<Sun and Sun<Des: 
		return False
		attrib=True
	if not vernalpointdaylight and Asc>Sun and Sun>Des: 
		return True
		attrib=True
	if not attrib and vernalpointdaylight:
		return True
	if not attrib and not vernalpointdaylight:
		return False

def transition(Sun,cusp):#should return true when the sun is in 1,12,6 or 7 (cusp is a list of all houses)
	#seems not to work well
	def isin(point, limit_a,limit_b):#limit_a should be authorised to be > limit_b and still check the small angle
		#is in is not thought in the logic of a circle: serious problems occur
			
		if limit_a>limit_b : vernal_point_in_range=True
		else:
			vernal_point_in_range=False
		attrib=False
		if vernal_point_in_range and limit_a<point and Sun<limit_b: 
			return False
			attrib=True
		if not vernal_point_in_range and limit_a>point and Sun>limit_b: 
			return True
			attrib=True
		if not attrib and vernal_point_in_range:
			return True
		if not attrib and not vernal_point_in_range:
			return False
	attrib_2=False
	if isin(Sun,cusp[0][11],cusp[0][1]):
		attrib_2=True
		print "dawn"
		return True
	if isin(Sun,cusp[0][7],cusp[0][8]):
		print"dusk"
		attrib_2=True
		return True	
	if not attrib_2: 
		print"day or night"
		return False

def search_value_without_space(charlist):	#works with any number of characters #unfinished
	global row_to_add
	print "len", len(charlist)
	def find_ref(wl,line):#to find the reference of the found match
		count=0
		char=0
		for j in range (20):
			#print wl[line-j]
			for i in range(len(wl[line-j])):#parse characters in line
				word_num=len(wl[line-j])-i-1
				line_num=line-j
				#print "line",line_num, "of lenght",len(wl[line]), "word",word_num
				word=wl[line_num][word_num]
				#print word
				for char in word:
					#print ord(char)
					if ord(char)==45 :
						count=count+1
						#print"found minus"
						if count==5:
							count=0
							#print "current line",wl[line-j]
							#print "preceding line",wl[line-j-1]
							return wl[line-j-1]
	def all_found(list_of_boolean):
		result=True
		for i in range (len(list_of_boolean)):
			#print list_of_boolean [i]
			if not list_of_boolean [i]:
				result=False
		#print
		return result	
	
	def hasNumbers(inputString):
		return any(char.isdigit() for char in inputString)
	
	def remove_chr(inputString,charlist):
		newstring=""
		for char in inputString:
			if not char in unicode(charlist):
				newstring=newstring+char
		return newstring
	
	def remove_vowel(string):
		newstring=""
		for char in string:
			if not ord(char) in vowel_codes_int:
				newstring=newstring+char
		return newstring
	
	charstringord=""
	#for char in range(len(charlist)): print charlist[char],unichr(int(charlist[char]))
	word_list = codecs.open("hebot.txt", 'r',encoding="utf-8")
	mem_list=[]
	match=[]
	matchline=[]
	matchref=[]
	found=[]
	temporary_sum=0
	for i in range(len(charlist)):
		found.append(False)
	num_lines = sum(1 for line in word_list)
	#print num_lines
	print "loading bible"
	word_list = codecs.open("hebot.txt", 'r',encoding="utf-8")
	for i in range(num_lines):#parsing thru lines in bible
		mem_list.append(word_list.readline())#in the list there are two levels: line and word
		#print mem_list[i]
	#grouping per verse
	verset_list=[]
	verset_ref=""
	verset_text=""
	
	for line in range(len(mem_list)):
		if not "-" in mem_list[line]:
			if hasNumbers( mem_list[line]):
				verset_ref= mem_list[line]
			else:
				verset_text=verset_text+ mem_list[line]

		if u'\u05c3' in mem_list[line]:#could be obtained also with unichr(1475)
			#print mem_list [line]
			#print "removing:",u'\n \u05b0־\u05c3'
			verset_list.append([remove_chr(verset_ref, u"\n"),remove_chr(verset_text, u"\n"),remove_vowel(remove_chr(verset_text,u'\n \u05b0־\u05c3'))])
			verset_ref=""
			verset_text=""
			verset_text_spaceless=""#not used
	#print "verset 2:",unicode(verset_list[1][1])
	#print "verset 2:",unicode(verset_list[1][2])
	#print "verset ref",verset_list[1][0]
	#print "verset pure", verset_list[1]
	
	count=0
	consec=0
	lastcount=0
	print "now searching in versets"
	for verset in verset_list:#cycling thru versets in bible
		#count +=1
		#if len(verset[2])==72:
			
			#print verset[0], len (verset[2]), "consec",consec,"lastcount",lastcount,"count",count, "spacing", math.fabs(lastcount-count)
		#	row_to_add[0]=verset[0]
		##	row_to_add[1]=verset[2]
		#	row_to_add[2]=consec
		#	row_to_add[3]=math.fabs(lastcount-count)
			#blind_add_row(row_to_add)
			#print  math.fabs(lastcount-count)
		#	if math.fabs(lastcount-count)<=1:
			#	consec+=1
		#		if consec>1:
			#		pass
			#		print "------------------------------------------------"
		#	else:
		#		consec=0
		#		pass
		#	lastcount=count
		#if "Exod. 14" in verset[0]:
		#	print verset[0], len (verset[2]),"\n", verset[2]
		
		
		
		#print "verset",verset[0]
		start=0
		offset=0
		currentlen=0
		string=""
		temporary_sum=0
	
		#print "now searching string value"
		
		while start<=len(verset[2])-1:#cycling thru chars
			#print
			
			offset=start+currentlen
			char=verset[2][offset]
			string=string+char
			temporary_sum=temporary_sum+numeric_value(char)
		
			if offset==len(verset[2])-1:
				start=start+1
				currentlen=-1
				temporary_sum=0
				string=""

			if temporary_sum>numeric_value(charlist):
				#print"rewinding to",start+1
				start=start+1
				currentlen=-1
				string=""
				temporary_sum=0
			
			
			if temporary_sum==numeric_value(charlist):# if numeric value is found
				
				#print "adding",string
				match.append(string)
				matchref.append(verset[0])	
				currentlen=-1
				start=start+1
				temporary_sum=0
				string=""
			currentlen=currentlen+1
			
	#output of matches in psalms
	count=0
	already_line=[]

	for word in range(len(match)):
		i=-1
		row_to_add=[]
		for value in empty_row:
			i+=1
			row_to_add.append(empty_row[i])
			
		if word<=len (match)-1:
			if "Pss." in matchref[word] or True:#condition canceled after data storing implementation
				
				if not match[word] in already_line :#output the same character suite only once
					count=count+1
					row_to_add[0]=count
					row_to_add[2]=unicode(match[word])# hebrew word in unicode characters
					row_to_add[1]=""#line number in bible txt
					if count==102:
						#print "len",len(match[word]),len(charlist)
						pass
								
								
					word_no_vowel=[]		
					for char in match[word]:
						if not ord(char) in vowel_codes_int:
							word_no_vowel.append(char)
							
							
					if len(word_no_vowel)==len(charlist):
						row_to_add[3]=True
					else:
						row_to_add[3]=False
												
					correct_order=True# voyels should be removed
					i=2
					for char in match[word]:
						#print "searching for in",str(ord(char)), charlist,(str(ord(char)) in charlist)
						if not( (ord(char) in vowel_codes_int)or (not str(ord(char)) in charlist)):
							if count==102:
								pass
								#print "char",char
							if ord(char)==int(charlist[i]):
								if count==102:
									pass
									#print "char compare True",ord(char),int(charlist[i])
							else: 
								if count==102:
									pass
									#print "char compare False",ord(char),int(charlist[i])
								correct_order=False
							i-=1
							if i<0:
								i=0
					if correct_order:
						pass
						#print "correct order in count", count	
					row_to_add[4]=correct_order
					
					row_to_add[6]= len(word_no_vowel)
					
					if unichr(1470) in match[word]:
						row_to_add[5]=True#composed word
					else:
						row_to_add[5]=False
						
						
					row_to_add[7]=matchref[word]#ref of verse
					blind_add_row(row_to_add)
					#print count,'match',word+1,unicode(match[word]),"line",matchline[word],"in verse",matchref[word][0],matchref[word][1]#detect duplicate with line number
					already_line.append(match[word])
		f=0
	print "recorded",count,"matches in bible"
	
			

def search_any_heb_combi(charlist):	#works with any number of characters 
	#this version works with vowel removal, but keeps spaces
	
	#print "len", len(charlist)
	#creating a list of double letters in searched string
	i=-1
	j=-1
	double_letters=[]
	dbl_lett_str=""
	def is_in(char, double_letters):
		for item in double_letters:
			if item[0]==char:
				return True
		return False
	for char1 in charlist:
		i+=1
		for char2 in charlist:
			j+=1
			if char2==char1 and not is_in(char2, double_letters) and not i==j:
				double_letters.append([char2,i,j])
		j=-1

	# creating a string of those letters
	for item in double_letters:
		dbl_lett_str=dbl_lett_str+unichr(int(item[0]))
		
	print "double letters",dbl_lett_str
	#print "list", double_letters
	
	def find_ref(wl,line):#to find the reference of the found match
		count=0
		char=0
		for j in range (20):
			#print wl[line-j]
			for i in range(len(wl[line-j])):#parse characters in line
				word_num=len(wl[line-j])-i-1
				line_num=line-j
				#print "line",line_num, "of lenght",len(wl[line]), "word",word_num
				word=wl[line_num][word_num]
				#print word
				for char in word:
					#print ord(char)
					if ord(char)==45 :
						count=count+1
						#print"found minus"
						if count==5:
							count=0
							#print "current line",wl[line-j]
							#print "preceding line",wl[line-j-1]
							return wl[line-j-1]
	def all_found(list_of_boolean):
		result=True
		for i in range (len(list_of_boolean)):
			#print list_of_boolean [i]
			if not list_of_boolean [i]:
				result=False
		#print
		return result
		
	def count(char,string):
		count=0
		for char_watch in string:
			if char == char_watch:
				count+=1
		return count
		
	
	def hasNumbers(inputString):
		return any(char.isdigit() for char in inputString)
	
	def remove_chr(inputString,charlist):
		newstring=""
		for char in inputString:
			if not char in unicode(charlist):
				newstring=newstring+char
		return newstring
	
	def remove_vowel(string):
		newstring=""
		for char in string:
			if not ord(char) in vowel_codes_int:
				newstring=newstring+char
		return newstring
		
		
	charstringord=""
	#for char in range(len(charlist)): print charlist[char],unichr(int(charlist[char]))
	word_list = codecs.open("hebot.txt", 'r',encoding="utf-8")
	mem_list=[]
	match=[]
	matchline=[]
	matchref=[]
	found=[]
	for i in range(len(charlist)):
		found.append(False)
	num_lines = sum(1 for line in word_list)
	
	
	
	#opening bible
	#print num_lines
	word_list = codecs.open("hebot.txt", 'r',encoding="utf-8")
	for i in range(num_lines):#parsing thru lines in bible
		mem_list.append(word_list.readline())#in the list there are two levels: line and word
		#print mem_list[i]
		
		
	print "sorting bible per verset"
	#sorting per verset
	verset_list=[]
	verset_ref=""
	verset_text=""
	
	for line in range(len(mem_list)):
		if not "-" in mem_list[line]:
			if hasNumbers( mem_list[line]):
				verset_ref= mem_list[line]
			else:
				verset_text=verset_text+ mem_list[line]

		if u'\u05c3' in mem_list[line]:#could be obtained also with unichr(1475)# this is the end of verset character
			#print mem_list [line]
			#print "removing:",u'\n \u05b0־\u05c3'
			verset_list.append([remove_chr(verset_ref, u"\n"),remove_chr(verset_text, u"\n"),remove_vowel(remove_chr(verset_text,u'\u05c3\n'))])
			verset_ref=""
			verset_text=""
			verset_text_spaceless=""#not used
	#print "verset 2:",unicode(verset_list[1][1])
	#print "verset 2:",unicode(verset_list[1][2])
	#print "verset ref",verset_list[1][0]
	#print "verset pure", verset_list[1]	
	
	print "searching in bible"
	k=-1
	for verset in verset_list:#cycling thru lines in bible
		#print "line"
		k+=1
		if "Pss. 23:2" in verset[0]:
			#print verset[0],k
			#print verset[2]
			pass
		line=verset[2].split()
		for word in range(len(line)):# cycling thru words
			for char in line[word]:#cycling thru chars
				for j in range(len(charlist)):#cycling thru searched chars
					if ord(char) == int(charlist[j]):
						#print "found", charlist[j]
						found[j]=True
				doubles_are_doubles=True#start off with false would be more logic
				for item in double_letters:
					if len(double_letters)>0 and all_found(found):
						#print line[word]
						#print count (unichr(int(item[0])),line[word])
						#print count(unichr(int(item[0])),line[word])<2
						#print
						pass
					if  count(unichr(int(item[0])),line[word])<2:#I should have counted the number of occurences and verified if the same number is found
						doubles_are_doubles=False#it means
						#print "double not found"
				if all_found(found) and doubles_are_doubles:# if all three characters are found
					#print "adding",line[word]
					match.append(line[word])
					matchref.append(verset[0])
			found=[]
			for i in range(len(charlist)):
				found.append(False)
				
	#--------------------------------------------------output of matches in psalms
	count=0
	already_line=[]

	for word in range(len(match)):
		i=-1
		row_to_add=[]
		for value in empty_row:
			i+=1
			row_to_add.append(empty_row[i])
			
		if word<=len (match)-1:
			if "Pss." in matchref[word] or True:#condition canceled after data storing implementation
				
				if not matchref[word] in already_line:#output the same line only once
					count=count+1
					row_to_add[0]=count
					row_to_add[2]=unicode(match[word])# hebrew word in unicode characters
					row_to_add[1]=""#line number in bible txt
					if count==102:
						#print "len",len(match[word]),len(charlist)
						pass		
								
					word_no_vowel=[]		
					for char in match[word]:
						if not ord(char) in vowel_codes_int:
							word_no_vowel.append(char)
							
							
					if len(word_no_vowel)==len(charlist):
						row_to_add[3]=True
					else:
						row_to_add[3]=False
												
					correct_order=True# voyels should be removed
					i=len(charlist)-1
					for char in match[word]:
						#print "searching for in",str(ord(char)), charlist,(str(ord(char)) in charlist)
						if not( (ord(char) in vowel_codes_int)or (not str(ord(char)) in charlist)):
							if count==102:
								pass
								#print "char",cha
							if ord(char)==int(charlist[i]):
								if count==102:
									pass
									#print "char compare True",ord(char),int(charlist[i])
							else: 
								if count==102:
									pass
									#print "char compare False",ord(char),int(charlist[i])
								correct_order=False
							i-=1
							if i<0:
								i=0
					if correct_order:
						pass
						#print "correct order in count", count	
					row_to_add[4]=correct_order
					
					row_to_add[6]= len(word_no_vowel)
					
					if unichr(1470) in match[word]:
						row_to_add[5]=True#composed word
					else:
						row_to_add[5]=False
						
						
					row_to_add[7]=matchref[word]#ref of verse
					blind_add_row(row_to_add)
					#print count,'match',word+1,unicode(match[word]),"line",matchline[word],"in verse",matchref[word][0],matchref[word][1]#detect duplicate with line number
					already_line.append(matchref[word])
		f=0
	print "recorded",count,"matches in bible"
	
	
	
def search_numeric_value_of(charlist):	#works with any number of characters #unfinished
	#only first mention of word
	only_first_match=True
	#print "len", len(charlist)
	def find_ref(wl,line):#to find the reference of the found match
		count=0
		char=0
		for j in range (20):
			#print wl[line-j]
			for i in range(len(wl[line-j])):#parse characters in line
				word_num=len(wl[line-j])-i-1
				line_num=line-j
				#print "line",line_num, "of lenght",len(wl[line]), "word",word_num
				word=wl[line_num][word_num]
				#print word
				for char in word:
					#print ord(char)
					if ord(char)==45 :
						count=count+1
						#print"found minus"
						if count==5:
							count=0
							#print "current line",wl[line-j]
							#print "preceding line",wl[line-j-1]
							return wl[line-j-1]
	def all_found(list_of_boolean):
		result=True
		for i in range (len(list_of_boolean)):
			#print list_of_boolean [i]
			if not list_of_boolean [i]:
				result=False
		#print
		return result	
	def no_vowel(word):
		#print "word", word
		word_no_vowel=[]		
		for char in word:
			if not ord(char) in vowel_codes_int:
				word_no_vowel.append(char)
		return word_no_vowel
	charstringord=""
	#--------------opening bible
	#for char in range(len(charlist)): print charlist[char],unichr(int(charlist[char]))
	word_list = codecs.open("hebot.txt", 'r',encoding="utf-8")
	mem_list=[]
	match=[]
	match_no_vowel=[]
	matchline=[]
	matchref=[]
	found=[]#unused
	for i in range(len(charlist)):
		found.append(False)
	num_lines = sum(1 for line in word_list)
	#print num_lines
	word_list = codecs.open("hebot.txt", 'r',encoding="utf-8")
	for i in range(num_lines):#parsing thru lines in bible
		mem_list.append(word_list.readline().split())#in the list there are two levels: line and word
		#print mem_list[i]
	for line in range(len(mem_list)):#cycling thru lines in bible
		#print "line"
		for word in range(len(mem_list[line])):# cycling thru words
			#converting word to charlist
			new_char_list=[]
			for char in no_vowel(mem_list[line][word]):
				new_char_list.append(str(ord(char)))
			if numeric_value(new_char_list)==numeric_value(charlist):
				if only_first_match:
					if not no_vowel(mem_list[line][word]) in match_no_vowel :
						match.append(mem_list[line][word])
				else:
					match.append(mem_list[line][word])
				match_no_vowel.append(no_vowel(mem_list[line][word]))
				matchline.append(line)
				matchref.append(find_ref(mem_list,line))
			
				
	#output of matches into a xls
	count=0
	already_line=[]

	for word in range(len(match)):
		i=-1
		row_to_add=[]
		for value in empty_row:
			i+=1
			row_to_add.append(empty_row[i])
			
		if word<=len (match)-1:
			if "Pss." in matchref[word][0] or True:#condition canceled after data storing implementation
				
				if not matchline[word] in already_line:#output the same line only once
					count=count+1
					row_to_add[0]=count
					row_to_add[2]=unicode(match[word])# hebrew word in unicode characters
					row_to_add[1]=matchline[word]#line number in bible txt
					if count==102:
						#print "len",len(match[word]),len(charlist)
						pass
								
					word_no_vowel=[]		
					for char in match[word]:
						if not ord(char) in vowel_codes_int:
							word_no_vowel.append(char)
							
							
					if len(word_no_vowel)==len(charlist):#boolean 1 of 3 same length
						row_to_add[3]=True
					else:
						row_to_add[3]=False
												
					correct_order=True# voyels should be removed
					i=2
					for char in match[word]:
						#print "searching for in",str(ord(char)), charlist,(str(ord(char)) in charlist)
						if not( (ord(char) in vowel_codes_int)or (not str(ord(char)) in charlist)):
							if count==102:
								pass
								#print "char",char
							if ord(char)==int(charlist[i]):
								if count==102:
									pass
									#print "char compare True",ord(char),int(charlist[i])
							else: 
								if count==102:
									pass
									#print "char compare False",ord(char),int(charlist[i])
								correct_order=False
							i-=1
							if i<0:
								i=0
					if correct_order:
						pass
						#print "correct order in count", count	
					row_to_add[4]=correct_order#boolean 2 of 3 correct order
					
					row_to_add[6]= len(word_no_vowel)
					
					if unichr(1470) in match[word]:
						row_to_add[5]=True#:#boolean 3 of 3 composed word
					else:
						row_to_add[5]=False
						
						
					row_to_add[7]=matchref[word][0],matchref[word][1]#ref of verse
					new_char_list=[]
					for char in match[word]:
						new_char_list.append(str(ord(char)))
					row_to_add[8]=numeric_value_list(new_char_list)
					blind_add_row(row_to_add)
					#print count,'match',word+1,unicode(match[word]),"line",matchline[word],"in verse",matchref[word][0],matchref[word][1]#detect duplicate with line number
					already_line.append(matchline[word])
		f=0
	print "recorded",count,"matches in bible"
	
	
	
def convert_to_charlist(string):
	charlist=[]
	for char_n in range(len(string)):
		#print ord(string[len(string)-char_n-1])
		charlist.append(str(ord(string[len(string)-char_n-1])))
	return charlist

def numeric_value(charlist):
	new_charlist=[]
	#print "type:",type(charlist)
	if not type(charlist)==list:
		new_charlist.append(str(ord(charlist)))
		charlist=new_charlist
		#print "changing:",charlist
	def find_offset(searched, list_i):
		i=0
		result=0
		#print "searched",searched, "in", list_i
		for i in range(len(list_i)):
			#print type (searched)
			#print type (list_i[i])
			if searched == list_i[i]:
				result=i
				#print "i=",i
				#no handling of mutiple match
				#last match offset is returned
	#	print i
		return result# 0 is returned if char is unfound
	
	#create the list of numeric values
	list_offset=[]
	list_values=[]
	for i in range (len(charlist)):
		list_offset.append(find_offset(charlist[i], char_codes))
	for i in range(len(list_offset)):
		list_values.append(num_values[list_offset[i]])
	#print "list values", list_values
	return sum(list_values)
		
def numeric_value_list(charlist):
	#print "len", len(charlist)
	def find_offset(searched, list_i):
		i=0
		result=0
		#print "searched",searched, "in", list_i
		for i in range(len(list_i)):
			#print type (searched)
			#print type (list_i[i])
			if searched == list_i[i]:
				result=i
				#print "i=",i
				#no handling of mutiple match
				#last match offset is returned
	#	print i
		return result# 0 is returned if char is unfound
	
	#create the list of numeric values
	list_offset=[]
	list_values=[]
	for i in range (len(charlist)):
		list_offset.append(find_offset(charlist[i], char_codes))
	for i in range(len(list_offset)):
		list_values.append(num_values[list_offset[i]])
	#print "list values", list_values
	string=""
	for value in list_values:
		if value<>0:
			string=string+str(value)+" "
	return string
		
		
def chain_search(method):
	global output_file
	save_output_file=output_file
	for i in range(21):
		word=Read_cell(1,i+1)
		print word
		print "Document generation started at ",datetime.datetime.strftime(datetime.datetime.now(),"%H:%M")
		charcodelist=convert_to_charlist(word)
		output_file=save_output_file.split(".")[0]+str(i).zfill(2)+"."+output_file.split(".")[1]
		#creating empty book
		data_book = {'Sheet1':[]}
		##pyexcel.save_as(array=data, dest_file_name=output_file)
		#sheet = pyexcel.Sheet(data)
		other_book = pyexcel.Book(data_book)
		other_book.save_as(output_file)
		
		search_any_heb_combi(charcodelist)#keeps spaces
		
		print "Document generation ended at ",datetime.datetime.strftime(datetime.datetime.now(),"%H:%M")

	output_file=save_output_file
		
#get values used for choosing the word------------------------------------------------------------------------

year= int (datetime.datetime.strftime(datetime.datetime.now(),"%Y"))
month= int (datetime.datetime.strftime(datetime.datetime.now(),"%m"))
day= int (datetime.datetime.strftime(datetime.datetime.now(),"%d"))
hour=int (datetime.datetime.strftime(datetime.datetime.now(),"%H"))
weekday=int (datetime.datetime.strftime(datetime.datetime.now(),"%w"))

print "year",year
divine_year=reduc_theo(year,2)
print "reduction theosophique de l'annee", divine_year
print "month",month
print "day",day
print "hour",hour
print "week day",weekday

#arrondir l'heure?

if not heure_dete(month,day) : 
	
	print "winter hour"
	uthour=hour-1
else: 	
	uthour=hour-2
	print "summer hour"
divine_year=reduc_theo(hour,2)
uthour=uthour+2
#print "uthour",uthour

current_jul_day_UT = swe.julday(year, month, day, uthour, gregflag)

ret_flag_sun = swe.calc_ut(current_jul_day_UT,0, gregflag); 
longitude_sun = ret_flag_sun[0]
print "Sun",longitude_sun
zodiaque=int(longitude_sun//30)+1
print "zodiaque",zodiaque


ret_flag_moon = swe.calc_ut(current_jul_day_UT,1, gregflag); 
longitude_moon = ret_flag_moon[0]
print "moon",longitude_moon

cusp = range(13)
cusp = swe.houses(current_jul_day_UT,geolat,geolon,"P")
#print cusp[0]
longitude_asc = cusp[0][0]
print "ascendant",longitude_asc

enjournee=daylight(longitude_sun, longitude_asc)
print "daylight",enjournee

print "-------------------------------------------------------------"
charcodelist=[]
charcodelistone=[]
#choose the letters of the word-------------------------------------------------------
#12 simples: zodiaque
vowel_codes=["05B0","05B1","05B2","05B3","05B4","05B5","05B6","05B7","05B8","05B9","05BB","05BC","05BD","05BF","05C1","05C2","05C4"]
vowel_codes_int=[]
for value in vowel_codes:
	vowel_codes_int.append(int(value,16))
# found in www.i18nguy.com/unicode/hebrew
#convert to int with int

Simplescodes=("1492","1493","1494","1495","1496", "1497","1500", "1504", "1505", "1506", "1510", "1511")
Simples=     ("He",  "Vav", "Zain","Heit","Teith","Yod", "Lamed","Noun","Sameck","Ayin","Tsade","Qoph")
Simples_Num=(5,6,7,8,9,10,30,50,60,70,90,100) #il manque l'octave des centaines
#selon une version de l'arbre de vie les doubles sont les diagonales de l'arbre.
# cette version serait ainsi celle la plus proche du sepher yetsirah
# la version de Dion fortune n'était pourtant pas mal.


#7 doubles: week
Doublescodes=("1499","1514","1491", "1512",  "1490", "1508", "1489")
Doubles=(     "Kaph","Tav","Daleth","Reish","Guimel","Phe","Beith")
Doubles_Num=(20,400,4,200,3,80,2)#il manque l'octave des centaines
#selon une version de l'arbre de vie les simples sont les piliers de l'arbre

#3 mothers: day night, dusk/dawn
Motherscodes=("1488","1502","1513")
Mothers=("Aleph","Mem","Shin")
Mothers_Num=(1,40,300) #il manque l'octave des milliers
#selon une version de l'arbre de vie les mères sont les équilibres de l'arbre


#finals:
finals=(      "final kaf", "final mem", "final nun","final_pe", "final tsade")
finals_codes=("05DA",      "05DD",     "05DF",      "05E3",    "05E5")
finals_codes_int_str=("1498",      "1501",     "1503",      "1507",    "1509")
final_num_values=(500,      600,       700,         800,       900)

#all numeric values
char_codes=("0000","1492","1493","1494","1495","1496", "1497","1500", "1504", "1505", "1506", "1510", "1511","1499","1514","1491", "1512",  "1490", "1508", "1489","1488","1502","1513","1498",      "1501",     "1503",      "1507",    "1509")
num_values=(0,5,6,7,8,9,10,30,50,60,70,90,100,20,400,4,200,3,80,2,1,40,300,500,      600,       700,         800,       900)

#punctuation
#maqaf tiret
#Sof Pasuq end of verset
punctuation=(      "Maqaf", "Paseq", "Sof Pasuq","Geresh", "Guershayim")
punctuation_codes=("05BE",   "05C0",  "05C3",      "05F3",    "05F4")
punctuation_codes_int=[]
for value in punctuation_codes:
	punctuation_codes_int.append(int(value,16))

print "punctuation:",punctuation_codes_int
finals_codes_int=[]
for value in finals_codes:
	finals_codes_int.append(int(value,16))
	
print finals_codes_int
print "**",Doubles[weekday] ,"**", "for week day (among 7 doubles)"
charcodelist.append(Doublescodes[weekday])
print "**",Simples[zodiaque-1],"**", "for sun sign (among 12 simples)"
charcodelist.append(Simplescodes[zodiaque-1])
transition_=transition(longitude_sun, cusp) 
transition_=False
enjournee=True
if transition_: 
	print "**",Mothers[0],"**", "because in dawn or dusk (problems with that one)"
	charcodelist.append(Motherscodes[0])
else:
	if enjournee :
		print "**",Mothers[2],"**","for hour sign (among 3 mothers) (day (shin))"
		charcodelist.append(Motherscodes[2])
	else:
		print "**",Mothers[1],"**","night (mem)"
		charcodelist.append(Motherscodes[1])
print "---------------------------------------------------------------"

#search_hebrew_word(("1499","1508","1488"))

#Mots=("Aleph","Qof","Guimel")("vase - aleph guimel reish mem lamed ","noyer - aleph guimel vav zain"
#qof-daleth-shin quadash holy
#qidqmti - qof dalteth mem tav iod - rise

#for char in range(len(charcodelist)): print charcodelist[char],unichr(int(charcodelist[char]))

#charcodelist=("1493","1512","1491")#rose

#charcodelist=["1493","1512","1491"]#rose

chain_search("spaceless")
	
#--------------------------------------------------------------------------------
#hebstring=u"אליהו"
#--------------------------------------------------------------------------------
#print "searched string length:",len(hebstring), "string:", hebstring
#charcodelist=convert_to_charlist(hebstring)
#print "num value of searched string", numeric_value(charcodelist)

#print charcodelist
#for char in range(len(charcodelist)): print charcodelist[char],unichr(int(charcodelist[char]))
#print "Document generation started at ",datetime.datetime.strftime(datetime.datetime.now(),"%H:%M")
#search_any_heb_combi(charcodelist)
#search_numeric_value_of(charcodelist)
#search_value_without_space(charcodelist)
print "Document generation ended at ",datetime.datetime.strftime(datetime.datetime.now(),"%H:%M")
print "---------------------------------------------------------------"
#print "searching aleph"
#charcodelistone.append("1488")
#search_hebrew_word(charcodelistone)

