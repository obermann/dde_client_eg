# dde_client_eg
Complete implementation of DDE (DDEML) Client as Python module.
Works with Python 2.6 and 3.x, tested on Windows 2000 and 10.

Read docstrings and use **DDEClient** class object that is like a dictionary of **DDEConversation** class ojects with all the functional methods, e.g.:

    dde_client = DDEClient()
    dde_client[("Service","Topic")].execute("This")
    print "There are %i ready conversations." % len(dde_client)

Look at my [XMPlay EvetGhost plugin](https://github.com/obermann/XMPlay) for how to use **dde_client_eg** for DDEML callback processing, *execute* and *request* methods, normal error handling.


pywin32 dde module emulator is not tested!
