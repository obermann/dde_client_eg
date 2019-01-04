# -*- coding: utf-8 -*-
#
"""
Service: type in Excel.

Topic: type the name of your worksheet file. 
This name is always exactly what is shown after the dash 
in the title-bar of the Excel spreadsheet. 
So if the title bar reads, "Microsoft Excel - Book1", 
then your Topic is simply Book1, but if the title bar 
reads Microsoft Excel - Test.xls then your Topic 
needs to be Test.xls.

[Note] 
If you want to link to a cell or range which is not 
on the first sheet in the workbook, you need to put 
the filename in square brackets, followed by the sheet name. 
For example, if your worksheet name is Test.xls:
The first sheet: Excel + Test.xls 
An unnamed sheet (e.g. Sheet2): Excel + [Test.xls]Sheet2 
A named sheet (e.g. StockData): Excel + [Test.xls]StockData 

"""
import win_unicode_console
win_unicode_console.enable(use_unicode_argv=True)

from ctypes import *
from ctypes.wintypes import *
user32 = WinDLL('user32', use_last_error=True)
import threading
import msvcrt
import codecs
from dde_client_eg import *



def unpack(x):
    assert isinstance(x, bytearray)
    if x[-2:] == b"\00\00":
        return x[:-2].decode("utf-16le")
    else:
        return codecs.encode(x, "hex") #[hex(b) for b in data]

        
        
def callback(type, data, **kwargs):
    print "Called back."
    if type == XTYP_XACT_COMPLETE:
        print "Transaction ID: ", kwargs["id"]
        if isinstance(data, bool):
            print "Success: ", data
        else:
            print unpack(data)
    elif type == XTYP_ADVDATA:
        if data:
            print unpack(data)
        else:
            print "XTYPF_NODATA"
        
        
        
excelfile = u"Ąžuolėlis.xls"

dde_client = DDEClient(callback=callback)

# "dummy" hack makes this conversation unique duplicate of ("ExCeL", excelfile)
retval = dde_client[("excel", excelfile, "dummy")].request(u"Skelbėjas", format=CF_UNICODETEXT)
print "Returned Sync UTF: ", codecs.encode(retval, "hex"), retval.decode("utf-16le")

retval = dde_client[("excel", excelfile)].request(u"Skelbėjas", format=CF_TEXT)
print "Returned Sync TXT: ", codecs.encode(retval, "hex"), retval.decode("latin-1")

retval = dde_client[("eXceL", excelfile)].request(u"Skelbėjas", timeout=TIMEOUT_ASYNC)
print "Returned Async: ", retval

retval = dde_client[("Excel", excelfile)].advise(u"Skelbėjas")
print "Returned Advise: ", retval

# wildconnect test
retval = dde_client[(None, excelfile)].poke(u"Gavėjas", u"Aš girdžiu liūtą!", timeout=TIMEOUT_ASYNC)
print "Returned Async Poke: ", retval
print

# findout who wildconnected
info = dde_client[(None, excelfile)].info(retval)
print "Conversation info of the transaction: ", retval
print "Service: ", info.SvcPartner 
print "Topic: ", info.Topic
print "Item: ", info.Item
print
print "Case insensitivity test: ", (u"excel", u"ąžuolėlis.xls") in dde_client
print
print "There are %i ready conversations." % len(dde_client)
print
print dde_client
print
print dde_client.__repr__()
print



def kbd_loop(ident):
    """Listen to keyboard."""
    WM_QUIT = 0x0012
    while True:
        if ord(msvcrt.getch()) == 27:
            user32.PostThreadMessageW(ident, WM_QUIT, 0, 0)
            break
            
def msg_loop():
    """Run the main windows message loop."""
    msg = MSG()
    while True:
        bRet = user32.GetMessageW(byref(msg), None, 0, 0)
        if not bRet:
            break
        if bRet == -1:
            raise WinError(get_last_error())
        user32.TranslateMessage(byref(msg))
        user32.DispatchMessageW(byref(msg))

if __name__ == '__main__':
    threading.Thread(target=kbd_loop, args=(threading.current_thread().ident,)).start()
    msg_loop()

    
# Finalize

dde_client[("Excel", excelfile)].advise(u"Skelbėjas", stop=True)
dde_client.shutdown()
print "Goodbye"
