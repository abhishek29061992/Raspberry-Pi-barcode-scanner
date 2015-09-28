import usb.core 
import usb.util
import time
import datetime
import ezodf #open office lib
import re
import linecache

#usb device specs (see usb descriptor)
USB_IF      = 0 # Interface
USB_TIMEOUT = 10000 # Timeout in MS
#barcode scanner used: TVS BSC 101
USB_VENDOR  = 0x24ea #device specific
USB_PRODUCT = 0x0197 #device specific

#character mapping:
index=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
index2=['1','2','3','4','5','6','7','8','9','0',40,'','','','','-']
code=""

#opening spreadsheet:
#spreadsheet=ezodf.newdoc(doctype='ods', filename='books.ods')
spreadsheet=ezodf.opendoc("book.ods")
sheets=spreadsheet.sheets
sheet=sheets['list']	#sheet named 'list'
endcol=sheet.ncols()	#no. of columns

if (endcol<4):
	for i in range (endcol,5):
		sheet.append_columns()	#append 5 cols. if<4
#columns:
sheet['A1'].set_value("S.No.")
sheet['B1'].set_value("ISBN")
sheet['C1'].set_value("Date Last Modified")
sheet['D1'].set_value("Qty.")

#append rows:
if sheet.nrows()==1:
	sheet.append_rows(1)
	end=2
else:
	if (sheet['B'+str(sheet.nrows())].value)==None :
		end=sheet.nrows()
	else:
		sheet.append_rows(1)
		end=sheet.nrows()

print (str(end)+" rows")
print (str(sheet.ncols())+" columns")
spreadsheet.save()


#initialise usb device:
dev = usb.core.find(idVendor=USB_VENDOR, idProduct=USB_PRODUCT)
endpoint = dev[0][(0,0)][0]

if dev.is_kernel_driver_active(USB_IF) is True: #detach kernel driver for device access by this app
  dev.detach_kernel_driver(USB_IF)

dev.set_configuration()
usb.util.claim_interface(dev, USB_IF)
print "Ready...\n"

#forvere loop; can be run using crontab at ppowerup
while True:
    control = None
    try:
     control = dev.read(endpoint.bEndpointAddress, endpoint.wMaxPacketSize, timeout=None,interface= USB_IF)
     arr=control
     if (arr[0]==2):
	char=index[(arr[2]-4)]
     elif (arr[0]==0):
	char=index2[(arr[2]-30)]

     if (char==40):
	print code
	sheet['B'+str(end)].set_value(str(code))
	sheet.append_rows(1)
	code=""
	sheet['C'+str(end)].set_value(str(datetime.datetime.now()))
	spreadsheet.save()
	end=end+1
	print "Added successfully"
     else:
	code+=str(char)

    except:
     pass
	#time.sleep(0.01) #Let CTRL+C actually exit

print "end"

