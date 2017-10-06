#Include ComImplementationBase.ahk
#Include UIAutomationClient_1_0_64bit.ahk
#SingleInstance Force



class UIA_SnoopWindowsDestroyed extends IUIAutomationEventHandlerImpl {
	;
	; 	Handles a Microsoft UI Automation event, called by the parent class which received the
	;		event call.
	;	@see: https://msdn.microsoft.com/en-us/library/windows/desktop/ee696045(v=vs.85).aspx
	;
	;	IUIAutomationEventHandler 	pInterface	Pointer to our interface address
	;	IUIAutomationElement		pSender		Pointer to Element for which Event happened
	;	EVENTID 					EventID 	The event identifier.
	;		@see: https://msdn.microsoft.com/en-us/library/windows/desktop/ee671223(v=vs.85).aspx
	;
	HandleAutomationEvent(pInterface, pSender, EventID) {
FileAppend, % "UIA_SnoopWindowsDestroyed `r`n", output.txt
		Sender := new IUIAutomationElement(pSender)
		hWnd := Sender.CurrentNativeWindowHandle
		WinGet, sProcess, ProcessName, ahk_id %hWnd%
		FileAppend % Format("Window Destroyed - hwnd={:X}, CurrentName={:20.20s}, Class={:20.20s}, Process={:s}", hWnd, Sender.CurrentName, Sender.CurrentClassName, sProcess) "`r`n", output.txt
	}
}

class UIA_SnoopWindowsCreated extends IUIAutomationEventHandlerImpl {
	;
	; 	Handles a Microsoft UI Automation event, called by the parent class which received the
	;		event call.
	;	@see: https://msdn.microsoft.com/en-us/library/windows/desktop/ee696045(v=vs.85).aspx
	;
	;	IUIAutomationEventHandler 	pInterface	Pointer to our interface address
	;	IUIAutomationElement		pSender		Pointer to Element for which Event happened
	;	EVENTID 					EventID 	The event identifier.
	;		@see: https://msdn.microsoft.com/en-us/library/windows/desktop/ee671223(v=vs.85).aspx
	;
	HandleAutomationEvent(pSender, EventID) {
FileAppend, % "UIA_SnoopWindowsCreated `r`n", output.txt
		Sender := new IUIAutomationElement(pSender)
		hWnd := Sender.CurrentNativeWindowHandle
		WinGet, sProcess, ProcessName, ahk_id %hWnd%
		FileAppend % Format("Window Created - hwnd={:X}, CurrentName={:20.20s}, Class={:20.20s}, Process={:s}", hWnd, Sender.CurrentName, Sender.CurrentClassName, sProcess) "`r`n", output.txt
	}
}

global OnCreated := new UIA_SnoopWindowsCreated()
global OnDestroyed := new UIA_SnoopWindowsDestroyed()
global UIA := CUIAutomation()

global DesktopElem := UIA.GetRootElement()

FileAppend % Format("AddAutomationEventHandler, DesktopElem={}, OnCreated.pInterface={:X}", DesktopElem.CurrentName(), OnCreated.pInterface ) "`r`n", output.txt

hr := UIA.AddAutomationEventHandler( UIAutomationClientTLConst.UIA_Window_WindowOpenedEventId
	, DesktopElem
	, UIAutomationClientTLConst.TreeScope_Children
	, 0
 	, OnCreated.pInterface )

	FileAppend % "hr= " Format("{:x}", hr & 0xFFFFFFFF) "`r`n", output.txt

UIA_Exit() {
	FileAppend % Format(A_ThisFunc "()") "`r`n", output.txt
	; Remove our exit function from the list
	OnExit(A_ThisFunc, 0)

	if (IsObject(UIA)) {
		UIA.RemoveAllEventHandlers()
		UIA := ""
	}

	if (DesktopElem) {
		ObjRelease(DesktopElem)
		DesktopElem := 0
	}

	if(OnCreated)
		OnCreated :=
	if(OnDestroyed)
		OnDestroyed :=

	return 0
}

OnExit("UIA_Exit")


F1::Exit, 0
