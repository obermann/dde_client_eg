# dde_client_eg
Complete implementation of DDE Client as Python module.
Works with Python 2.6 and 3.x, tested on Windows 2000 and 10.

Read docstrings and use **DDEClient** class object that is like a dictionary of **DDEConversation** class ojects with all the functional methods, e.g.:

    dde_client = DDEClient()
    dde_client[("Service","Topic")].execute("This")
    print "There are %i ready conversations." % len(dde_client)

pywin32 emulator is not tested!
