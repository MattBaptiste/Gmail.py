import sqlite3
import os
import shutil
import xlsxwriter
import sys
import time


#log file
log = open('EmailLogFile.txt', 'w')
log.write("start " + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
log.write('\n')


emaildb = "data\\data\\com.android.email\\databases\\EmailProvider.db"
log.write("enters drive location" + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
log.write('\n')


#opens database
db0 = sqlite3.connect(emaildb)
curser0 = db0.cursor()
curser0.execute('''SELECT _id FROM account''')
all_rows = curser0.fetchall()

#sets max number users
for row in all_rows:
	maxuser = row[0]

#close database and courser
curser0.close()
db0.close()
log.write("set total amount of users" + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
log.write('\n')

while maxuser >= 1:
	bookname = "user" + repr(maxuser) + ".xlsx"
	
	
	#name work book 
	book = xlsxwriter.Workbook(bookname)
	sheet = book.add_worksheet('overview')
	sheet2 = book.add_worksheet('logon info')
	sheet3 = book.add_worksheet('settings')
	sheet4 = book.add_worksheet('attachments')
	format = book.add_format()
	format.set_pattern(1)
	format.set_bg_color('cyan')
	format2 = book.add_format({'bold' : True, 'bg_color' : 'silver'})

	
	log.write("built excel book for out put" + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	
	#sets row and column counter
	_row = 0
	_col = 0

	#over view of user 
	db = sqlite3.connect(emaildb)
	curser = db.cursor()
	curser.execute('''SELECT _id, toList, fromList, timeStamp, subject, snippet,ccList, bccList FROM message WHERE accountKey =?''',(maxuser,))
	all_rows = curser.fetchall()
	
	log.write("start writing overview for " + bookname + " " + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	#tiles of columns	
	if _row % 2 == 0:
		sheet.write(_row, _col, 'id', format2) #0
		sheet.write(_row, _col + 1, 'To Name', format2) #1
		sheet.write(_row, _col + 2, 'To Address', format2) #2
		sheet.write(_row, _col + 3, 'From Name', format2) #3
		sheet.write(_row, _col + 4, 'From Address', format2)#4
		sheet.write(_row, _col + 5, 'Time Sent', format2) #5
		sheet.write(_row, _col + 6, 'Subject', format2) #6
		sheet.write(_row, _col + 7, 'Snippet', format2)#7
		sheet.write(_row, _col + 8, 'cc Address', format2)#8
		sheet.write(_row, _col + 9, 'bcc Addrerss', format2)#9
		_row += 1
	else:
		sheet.write(_row, _col, 'id') #0
		sheet.write(_row, _col + 1, 'To Name') #1
		sheet.write(_row, _col + 2, 'To Address') #2
		sheet.write(_row, _col + 3, 'From Name') #3
		sheet.write(_row, _col + 4, 'From Address')#4
		sheet.write(_row, _col + 5, 'Time Sent') #5
		sheet.write(_row, _col + 6, 'Subject') #6
		sheet.write(_row, _col + 7, 'Snippet')#7
		sheet.write(_row, _col + 8, 'cc Address')#8
		sheet.write(_row, _col + 9, 'bcc Addrerss')#9
		_row += 1
	sheet.freeze_panes(1,0)
	
	# writing called info to rows
	for row in all_rows:
		if _row % 2 == 0:
			sheet.write(_row, _col, row[0], format)
			if '<' in row[1]:
				address = row[1]
				addresslist = address.split('<')
				addresslist[1] = addresslist[1].replace('>', "")
				addresslist[0] = addresslist[0].replace('"', "")
				sheet.write(_row, _col + 1, addresslist[0], format)
				sheet.write(_row, _col + 2, addresslist[1], format)
			else:
				addresslist = row[1]
				addresslist = addresslist.replace('>', "")
				addresslist = addresslist.replace('<', "")
				sheet.write(_row, _col + 1, "", format)
				sheet.write(_row, _col + 2, addresslist, format)
			if '<' in row[2]:
				address2 = row[2]
				address2list = address2.split('<')
				address2list[1] = address2list[1].replace('>', "")
				sheet.write(_row, _col + 3, address2list[0], format)
				sheet.write(_row, _col + 4, address2list[1], format)
			else:
				addresslist = row[2]
				addresslist = addresslist.replace('>', "")
				addresslist = addresslist.replace('<', "")
				sheet.write(_row, _col + 3, "", format)
				sheet.write(_row, _col + 4, addresslist, format)
			sheet.write(_row, _col + 5, time.strftime("%m/%d/%y %H:%M:%S", time.localtime(row[3]/1000)), format)
			sheet.write(_row, _col + 6, row[4], format)
			sheet.write(_row, _col + 7, row[5], format)
			if row[6]:
				sheet.write(_row, _col + 8, row[6], format)
			else: 
				sheet.write(_row, _col + 8, 'none', format)
			if row[7]:
				sheet.write(_row, _col + 9, row[7], format)
			else: 
				sheet.write(_row, _col + 9, 'none', format)
			_row +=1 	
		else:
			
			sheet.write(_row, _col, row[0])
			if '<' in row[1]:
				address = row[1]
				addresslist = address.split('<')
				addresslist[1] = addresslist[1].replace('>', "")
				addresslist[0] = addresslist[0].replace('"', "")
				sheet.write(_row, _col + 1, addresslist[0])
				sheet.write(_row, _col + 2, addresslist[1])
			else:
				addresslist = row[1]
				addresslist = addresslist.replace('>', "")
				addresslist = addresslist.replace('<', "")
				sheet.write(_row, _col + 1, "")
				sheet.write(_row, _col + 2, addresslist)
			if '<' in row[2]:
				address2 = row[2]
				address2list = address2.split('<')
				address2list[1] = address2list[1].replace('>', "")
				sheet.write(_row, _col + 3, address2list[0])
				sheet.write(_row, _col + 4, address2list[1])
			else:
				address2list = row[2]
				address2list = address2list.replace('>', "")
				address2list = address2list.replace('<', "")
				sheet.write(_row, _col + 3, "")
				sheet.write(_row, _col + 4, addresslist)
			sheet.write(_row, _col + 5, time.strftime("%m/%d/%y %H:%M:%S", time.localtime(row[3]/1000)))
			sheet.write(_row, _col + 6, row[4])
			sheet.write(_row, _col + 7, row[5])
			if row[6]:
				sheet.write(_row, _col + 8, row[6])
			else: 
				sheet.write(_row, _col + 8, 'none')
			if row[7]:
				sheet.write(_row, _col + 9, row[7])
			else: 
				sheet.write(_row, _col + 9, 'none')
			_row +=1 
		
	log.write("finished writing overview for " + bookname + " " + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	
	log.write("start writing account information" + bookname + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	
	#reset row and columns
	_row = 0
	_col = 0
	
	curser.execute('''SELECT _id, displayName, emailAddress, syncInterval, senderName, signature FROM account WHERE _id = ?''', (maxuser,))
	all_rows3 = curser.fetchall()
	if _row % 2 == 0:
		sheet3.write(_row, _col, 'id', format2)
		sheet3.write(_row, _col + 1, 'display name', format2)
		sheet3.write(_row, _col + 2, 'email address', format2)
		sheet3.write(_row, _col + 3, 'sync interval in minutes', format2)
		sheet3.write(_row, _col + 4, 'sender name', format2)
		sheet3.write(_row, _col + 5, 'signature', format2)
		_row += 1
	else:
		sheet3.write(_row, _col, 'id')
		sheet3.write(_row, _col + 1, 'display name')
		sheet3.write(_row, _col + 2, 'email address')
		sheet3.write(_row, _col + 3, 'sync interval in minutes')
		sheet3.write(_row, _col + 4, 'sender name')
		sheet3.write(_row, _col + 5, 'signature')
		_row += 1
	sheet3.freeze_panes(1,0) 
	for row in all_rows3:
		if _row % 2 == 0:
			sheet3.write(_row, _col, row[0], format)
			sheet3.write(_row, _col + 1, row[1], format)
			sheet3.write(_row, _col + 2, row[2], format)
			email = row[2]
			sheet3.write(_row, _col + 3, row[3], format)
			sheet3.write(_row, _col + 4, row[4], format)
			sheet3.write(_row, _col + 5, row[5], format)
			_row += 1
		else:
			sheet3.write(_row, _col, row[0])
			sheet3.write(_row, _col + 1, row[1])
			sheet3.write(_row, _col + 2, row[2])
			email = row[2]
			sheet3.write(_row, _col + 3, row[3])
			sheet3.write(_row, _col + 4, row[4])
			sheet3.write(_row, _col + 5, row[5])
			_row += 1
		
		
	log.write("finish writing account information" + bookname + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	
	log.write("start writing login information" + bookname + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	
	#reset row and columns
	_row = 0
	_col = 0	

	curser.execute('''SELECT login, password FROM hostauth WHERE login = ?''', (email,))
	all_rows2 = curser.fetchall()

	if _row % 2 == 0:
		sheet2.write(_row, _col, 'login', format2)
		sheet2.write(_row, _col + 1, 'name', format2)
		_row+=1
	else:
		sheet2.write(_row, _col, 'login')
		sheet2.write(_row, _col + 1, 'name')
		_row+=1
	sheet2.freeze_panes(1,0)
	for row in all_rows2:
		if _row % 2 == 0:
			sheet2.write(_row, _col, row[0], format)
			sheet2.write(_row, _col+ 1, row[1], format)
			_row +=1
		else:
			sheet2.write(_row, _col, row[0])
			sheet2.write(_row, _col+ 1, row[1])
			_row +=1
	
	log.write("end writing login information for" + bookname + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	
	curser2 = db.cursor()
	curser2.execute('''SELECT messagekey, fileName FROM attachment WHERE accountKey =?''',(maxuser,))
	all_rows2 = curser2.fetchall()
	
	log.write("start writing attachment information" + bookname + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	
	#reset rows and column
	_row = 0
	_col = 0
	
	#titles
	if _row % 2 == 0:
		sheet4.write(_row, _col, 'message number', format2)
		sheet4.write(_row, _col + 1, 'file name', format2)
		_row += 1
	
	else:
		sheet3.write(_row, _col, 'message number')
		sheet3.write(_row, _col + 1, 'file name')
		_row += 1
	sheet4.freeze_panes(1,0)
	for row in all_rows2:
		if _row % 2 == 0:
			sheet4.write(_row, _col, row[0], format)
			sheet4.write(_row, _col + 1, row[1], format)
			_row += 1
		else:
			sheet4.write(_row, _col, row[0])
			sheet4.write(_row, _col + 1, row[1])
			_row += 1
	
	
	log.write("f writing attachment information" + bookname + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	
	
	maxuser -= 1
	
	book.close()
	log.write("close book for " + bookname + ' '  + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	#close database and excel book
db.close()
log.write("close book" + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
log.write('\n')

#copy function 
point = "data\\data\\com.android.email\\files\\body"

shutil.copytree(point, 'emailexport')
log.write("copied out emails" + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))

point = "data\\data\\com.android.email\\shared_prefs"

shutil.copytree(point, 'shared_prefsexport')
log.write("copied out emails " + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
log.write('\n')
log.write("complete  " + time.strftime("%m/%d/%y %H:%M",time.localtime()))
