class myEvents

	private eventListeners
	'Dictionary object containing EventName, ListenersDictionary couples

	
	Private Sub Class_Initialize()
		set eventListeners = CreateObject ("Scripting.Dictionary")
	end sub
	
	public function addListener(eventName, functionRef)
		'If an EventName doesn't exists it creates it and adds the listener function reference
		if not eventListeners.exists(eventName) then
			eventListeners.add eventName, CreateObject("Scripting.Dictionary")
			eventListeners.item(eventName).add 1, functionRef
		'otherwise it adds the function reference to another dictionary entry on the same event
		else
			eventListeners.item(eventName).add (eventListeners.item(eventName).count + 1), functionRef
		end if		
	end function
	
	public function removelisteners(eventName)
		'removes all listeners from EventName
		eventListeners.item(EventName) = nothing
	end function
	
	Public Function raise(eventName)
		'raises the named EventName event, firing all the referenced listener functions
		for i = 1 to eventListeners.item(eventName).count
			eventListeners.item(eventName).item(i)()
		next
	end function

end class

