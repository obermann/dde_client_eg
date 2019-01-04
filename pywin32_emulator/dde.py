# -*- coding: utf-8 -*-
# 
"""
Drop-in pywin32 dde module emulator for DDE clients.
Exceptions that you may have to except are DDEError.
Default format for all transactions is CF_UNICODETEXT.

"""

import weakref
from dde_client_eg import (DDEClient, DDEError,
    APPCLASS_MONITOR,
    APPCLASS_STANDARD,
    APPCMD_CLIENTONLY,
    APPCMD_FILTERINITS,
    CBF_FAIL_ADVISES,
    CBF_FAIL_ALLSVRXACTIONS,
    CBF_FAIL_CONNECTIONS,
    CBF_FAIL_EXECUTES,
    CBF_FAIL_POKES,
    CBF_FAIL_REQUESTS,
    CBF_FAIL_SELFCONNECTIONS,
    CBF_SKIP_ALLNOTIFICATIONS,
    CBF_SKIP_CONNECT_CONFIRMS,
    CBF_SKIP_DISCONNECTS,
    CBF_SKIP_REGISTRATIONS,
    MF_CALLBACKS,
    MF_CONV,
    MF_ERRORS,
    MF_HSZ_INFO,
    MF_LINKS,
    MF_POSTMSGS,
    MF_SENDMSGS,
    CF_TEXT
)



class CreateServer:

    def __init__(self):
        self.dde_client = None
    
    def Create(self, *args): # ignored flags go here
        self.dde_client = DDEClient(flags=APPCMD_CLIENTONLY|CBF_SKIP_REGISTRATIONS)
        
    def Shutdown(self):
        dde_client.shutdown()
        self.dde_client = None
        
    def __del__(self):
        if self.dde_client:
            self.Shutdown()
        
    def Destroy(self):
        self.__del__()

    def GetLastError(self):
        return 0
        
        
    
class CreateConversation:
    
    NOTHING = object()
    DDE_TIMEOUT = 60000 # a minute instead of 5000ms?!
    
    def __init__(self, server):
        self._key = self.NOTHING
        assert isinstance(server.dde_client, DDEClient)
        self.dde_client = weakref.ref(server.dde_client)
        
    def __del__(self):
        if self._key != self.NOTHING:
            # here assuming a strict pattern
            # when one CreateConversation is used
            # with single ConnectTo(), but messy
            # pattern works well too with final
            # CreateServer Shutdown()
            del self.dde_client[self._key]
        self.dde_client = None
        
    def ConnectTo(self, service, topic):
        self._key = (service, topic)
        self.dde_client[self._key]
    
    def Connected(self):
        return False if self._key == self.NOTHING else self._key in self.dde_client

    def Exec(self, data):
        if self._key != self.NOTHING:
            self.dde_client[self._key].execute(data, timeout=self.DDE_TIMEOUT)

    def Poke(self, item, data):
        if self._key != self.NOTHING:
            self.dde_client[self._key].poke(item, data, timeout=self.DDE_TIMEOUT)

    def Request(self, item):
        if self._key != self.NOTHING:
            retval = self.dde_client[self._key].request(item, timeout=self.DDE_TIMEOUT)
            if retval[-2:] == b"\00\00":
                retval = retval[:-2]
            return retval.decode("utf-16le")
            # for messed up (non-unicode) DDE servers try instead:
            # retval = self.dde_client[self._key].request(item, format=CF_TEXT, timeout=self.DDE_TIMEOUT)
            # if retval[-1:] == b"\00":
                # retval = retval[:-1]
            # return retval.decode("latin-1")
