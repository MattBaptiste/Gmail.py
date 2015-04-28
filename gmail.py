import sqlite3
import os
import xlsxwriter
import sys
import zlib
import time
import shutil
import glob

#log file
log = open('GmailLogFile.txt', 'w')
log.write("start " + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
log.write('\n')
#copies mailstore database
mailstorepath = "data\\data\\com.google.android.gm\\databases\\mailstore.*.db"
files1 = glob.glob(mailstorepath)
dst = os.getcwd()
for file in files1:
	shutil.copy(file,dst)
log.write("copied mailstore databasess" + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
log.write('\n')

#copies internal database
internalpath = 	"data\\data\\com.google.android.gm\\databases\\internal.*.db"
files2 = glob.glob(internalpath)
for file in files2:
	shutil.copy(file,dst)
log.write("copied internal database" + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
log.write('\n')
	
i = 0
email = glob.glob('mailstore.*.db')
while i <= len(email) - 1: 
	email[i] = email[i].replace("mailstore.", "")
	i += 1

i = 0
while i <= len(email) - 1: 
	email[i] = email[i].replace(".db", "")
	i += 1



d = 0
i = 0
while d < len(email):

	dbemail = 'mailstore.' + email[i] + '.db'
	dbsettings =  'internal.' + email[i] + '.db'

	#name of workbook and worksheets
	book = xlsxwriter.Workbook(email[i] +'.xlsx')
	sheet = book.add_worksheet('overview')
	sheet2 = book.add_worksheet('Settings')
	sheet3 = book.add_worksheet('Attachments')
	format = book.add_format()
	format.set_pattern(1)
	format.set_bg_color('cyan')
	format2 = book.add_format({'bold' : True, 'bg_color' : 'silver'})
	
	log.write("opened excel files" + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	

	#sets row and column counter
	_row = 0
	_col = 0

	#calls database makes curser and fetches all rows
	db = sqlite3.connect(dbemail)
	curser = db.cursor()
	curser.execute('''SELECT _id, fromAddress, toAddresses, ccAddresses, bccAddresses, dateSentMs, dateReceivedMs, subject, snippet, bodyCompressed, messageId FROM messages''')
	all_rows = curser.fetchall()
	
	log.write("start writing to excel files for " + email[i] + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	log.write("start Overview for " + email[i] + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	#tiles of columns
	if _row % 2 == 0:
		sheet.write(_row, _col, 'id', format2) #0
		sheet.write(_row, _col + 1, 'From Name', format2) #1
		sheet.write(_row, _col + 2, 'From Address', format2) #2
		sheet.write(_row, _col + 3, 'To Address', format2) #3
		sheet.write(_row, _col + 4, 'To Address', format2) #4
		sheet.write(_row, _col + 5, 'cc Address', format2) #5
		sheet.write(_row, _col + 6, 'bcc Address', format2) #6
		sheet.write(_row, _col + 7, 'Time Sent', format2) #7
		sheet.write(_row, _col + 8, 'Time Received', format2)#8
		sheet.write(_row, _col + 9, 'Subject', format2) #9
		sheet.write(_row, _col + 10, 'Snippet', format2)#10
		sheet.write(_row, _col + 11, 'Message id', format2)#11
		_row += 1
	else:
		sheet.write(_row, _col, 'id') #0
		sheet.write(_row, _col + 1, 'From Name') #1
		sheet.write(_row, _col + 2, 'From Address') #2
		sheet.write(_row, _col + 3, 'To Name') #3
		sheet.write(_row, _col + 4, 'To Address') #4
		sheet.write(_row, _col + 5, 'cc Address') #5
		sheet.write(_row, _col + 6, 'bcc Address') #6
		sheet.write(_row, _col + 7, 'Time Sent') #7
		sheet.write(_row, _col + 8, 'Time Received')#8
		sheet.write(_row, _col + 9, 'Subject') #9
		sheet.write(_row, _col + 10, 'Snippet')#10
		sheet.write(_row, _col + 11, 'Message id')#11
		_row += 1
	sheet.freeze_panes(1,0)
	# writing called info to rows
	for row in all_rows:
		if _row % 2 == 0:
			sheet.write(_row, _col, row[0], format)
			word = row[1].split('<')
			name = word[0]
			name = name.replace('"', "")
			address = word[1]
			address = address.replace('>', "")
			sheet.write(_row, _col + 1, name, format)
			sheet.write(_row, _col + 2, address, format)
			word2 = row[2].split('<')
			toname = word2[0]
			toname = toname.replace('"', "")
			if len(word2) == 2:
				toaddress = word2[1]
				toaddress = toaddress.replace('>', "")
			else:
				toddress = ""
			sheet.write(_row, _col + 3, toname, format)
			sheet.write(_row, _col + 4, toaddress, format)
			sheet.write(_row, _col + 5, row[3], format)
			sheet.write(_row, _col + 6, row[4], format)
			sheet.write(_row, _col + 7, time.strftime("%m/%d/%y %H:%M:%S", time.localtime(row[5]/1000)), format)
			sheet.write(_row, _col + 8, time.strftime("%m/%d/%y %H:%M:%S", time.localtime(row[6]/1000)), format)
			sheet.write(_row, _col + 9, row[7], format)
			sheet.write(_row, _col + 10, row[8], format)
			sheet.write(_row, _col + 11, row[10], format)
			mid = row[10]
			_row +=1 
	
			name = repr(row[0])
			try:
				if row[9]:
					text = bytes.decode(zlib.decompress(row[9]))
					outbody = open(email[i] + ' ' + name+'.html', 'w')
					outbody.write('<html><body>' + text + '</body></html>')
			except UnicodeEncodeError:
				text = text.encode('ascii', 'ignore').decode('ascii')
				outbody = open(email[i] + ' ' + name+'.html', 'w')
				outbody.write('<html><body>' + text + '</body></html>')
				
		else:
			sheet.write(_row, _col, row[0])
			word = row[1].split('<')
			name = word[0]
			name = name.replace('"', "")
			address = word[1]
			address = address.replace('>', "")
			sheet.write(_row, _col + 1, name)
			sheet.write(_row, _col + 2, address)
			word2 = row[2].split('<')
			toname = word[0]
			toname = toname.replace('"', "")
			if len(word2) == 2:
				toaddress = word2[1]
				toaddress = toaddress.replace('>', "")
			else:
				toaddrress = ""
			sheet.write(_row, _col + 3, toname)
			sheet.write(_row, _col + 4, toaddress)
			sheet.write(_row, _col + 5, row[3])
			sheet.write(_row, _col + 6, row[4])
			sheet.write(_row, _col + 7, time.strftime("%m/%d/%y %H:%M:%S", time.localtime(row[5]/1000)))
			sheet.write(_row, _col + 8, time.strftime("%m/%d/%y %H:%M:%S", time.localtime(row[6]/1000)))
			sheet.write(_row, _col + 9, row[7])
			sheet.write(_row, _col + 10, row[8])
			sheet.write(_row, _col + 11, row[10])
			mid = row[10]
			
			_row +=1 
	
			name = repr(row[0])
			try:
				if row[9]:
					text = bytes.decode(zlib.decompress(row[9]))
					outbody = open(email[i] + ' ' + name+'.html', 'w')
					outbody.write('<html><body>' + text + '</body></html>')
			except UnicodeEncodeError:
				text = text.encode('ascii', 'ignore').decode('ascii')
				outbody = open(email[i] + ' ' + name+'.html', 'w')
				outbody.write('<html><body>' + text + '</body></html>')
				
	log.write("finished writing emails for " + email[i] + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	log.write("finished writing overview for " + email[i] + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	
	curser3 = db.cursor()
	curser3.execute('''SELECT messages_messageId, filename FROM attachments''')
	all_rows3 = curser3.fetchall()
	
	log.write("start writing attachments for " + email[i] + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	#reset row and column
	_row = 0
	_col = 0
	
	#titles
	if _row % 2 ==0:
		sheet3.write(_row, _col, 'message id', format2)
		sheet3.write(_row, _col + 1, 'filelocation', format2)
		_row += 1
	else:
		sheet3.write(_row, _col, 'message id')
		sheet3.write(_row, _col + 1, 'filelocation')
		_row += 1
	sheet3.freeze_panes(1,0)
	for row in all_rows3:
		if  _row % 2 ==0:
			sheet3.write(_row, _col, row[0], format)
			sheet3.write(_row, _col + 1, row[1], format)
			_row += 1
		else:
			sheet3.write(_row, _col, row[0])
			sheet3.write(_row, _col + 1, row[1])
			_row += 1



	
		
	log.write("finished writing attachments for " + email[i] + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	


	#closes database
	curser.close()
	db.close()

	#reset row and columns
	_row = 0
	_col = 0

	#opens settings database
	db2 =  sqlite3.connect(dbsettings)
	curser2 = db2.cursor()
	curser2.execute('''SELECT name, value FROM internal_sync_settings''')
	all_rows2 = curser2.fetchall()

	log.write("start writing settings for " + email[i] + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	#titles
	if _row % 2 ==0:
		sheet2.write(_row, _col, 'name', format2)
		sheet2.write(_row, _col + 1, 'value', format2)
		_row += 1
	else:
		sheet2.write(_row, _col, 'name')
		sheet2.write(_row, _col + 1, 'value')
		_row += 1
	sheet2.freeze_panes(1,0)
	for row in all_rows2:
		if _row % 2 ==0:
			sheet2.write(_row, _col, row[0], format)
			sheet2.write(_row, _col + 1, row[1], format)
			_row += 1
		else:
			sheet2.write(_row, _col, row[0])
			sheet2.write(_row, _col + 1, row[1])
			_row += 1
	
	log.write("finished writing settings for " + email[i] + ' ' + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
	log.write('\n')
	
	#closes internal settings database
	curser2.close()
	db2.close()
	#closes book 
	book.close()
	

	
	d += 1
	i += 1
	
#attachments 
log.write("start coying attachments " + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
log.write('\n')
location  = "data\\data\\com.google.android.gm\\cache"
shutil.copytree(location, 'GmailAttachments')
log.write("finish copying attachments, only has attachments for last user used on phone " + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
log.write('\n')

log.write("start writing xml files " + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
log.write('\n')
#copy function
scr = "data\\data\\com.google.android.gm\\shared_prefs"
#export = os.getcwd() + "export"
shutil.copytree(scr, 'export')
log.write("finished writing xml files " + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))
log.write('\n')
log.write("Complete " + time.strftime("%m/%d/%y %H:%M:%S",time.localtime()))

