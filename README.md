# VBS-Events
Adds basic event management capabilities to VBScript projects

Usage:

1. Create myEvents Instance
```vbscript
dim myEv
set myEv = new myEvents
```

2. Have some global scope functions/subs to manage events
```vbscript
function TestEventHandler()
    REM your code here
end function
```

3. Register the functions as listeners for a specific custom event
```vbscript
myEv.addListener "TestEvent" getRef("TestEventHandler")
```

4. Fire the event and enjoy!
```vbscript
myEv.raise "TestEvent"
```

Some notes:
- Event handlers functions/subs MUST be in global scope
- addListener and raise methods can be called from anywhere, even inside classes
