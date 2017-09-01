# TypeLib2AHK #

TypeLib2AHK is a free, open source tool to convert information about COM interfaces stored in type libraries to code usable with AutoHotkey 1.1 and AutoHotkey v2 (https://autohotkey.com/). 

The aim is to make it more comfortable and seamless to work with COM from AutoHotkey by providing class wrappers for structures and non-dispatch interfaces which are not natively supported by AutoHotkey and also the named constants associated with the interfaces.

## How to use ##

TypeLib2AHK.ahk can either be run directly as a standalone application or be used in an AutoHotkey script to retrieve information about specific COM objects (see examples at the beginning of the code in TypeLib2AHK.ahk). It can be used with AutoHotkey 32bit and 64bit Unicode releases. AutoHotkey ANSI releases are not supported.

TypeLib2AHK.ahk currently only runs under AutoHotkey 1.1 but can create code for AutoHotkey 1.1 and AutoHotkey v2.

The standalone application offers the user a list of type libraries stored on the computer and retrieved from applications currently in memory. Type library information can also be loaded from DLL files.

Type libraries can be viewed or directly converted. The viewing functionality is fairly basic; to get more complete information you can use Oleview.exe which is part of the Windows SDK, which can be downloaded for free from Microsoft.

## How to use the created code ##

The resulting code is structured as follows (all examples refer to the UIAutomationClient type library):

* Header: Information about the type library.
* CoClasses: The code in this section allows to instantiate the underlying COM interfaces. Example: UIA:=CUIAutomation()
* Alias type definitions: For information only as type definitions are usually not relevant in AHK.
* Constants (enumerations and modules): They can be used directly (Example: myTreeScope:=UIAutomationClientTLConst.TreeScope_Ancestors) or the wrapper class can be instantiated as needed. Reverse lookup of a constant name is also possible (Example ValueName:=UIAutomationClientTLConst.UIA_ControlTypeIds(Value))
* Structures (records and unions): The structures are wrapped as classes for transparent use in AHK scripts (see Example code below). Wrapped structures can be used directly as parameters in function calls or wrapped interfaces. 
* Interfaces: The interfaces are wrapped as classes for transparent use in AHK scripts (see Example code below).
* Dispatch interfaces: Dispatch interfaces are handled natively by AHK. Their contents are included for information only.

The code may not contain all the above sections, depending on what is defined in the respective type library.

IMPORTANT: 
* Before using the created code carefully read the definitions and make sure that it doesn't redefine any Autohotkey functions. As an example: mscorlib.dll overrides Object().
* Many basic structures are defined in several type libraries. You may need to edit the created code to avoid doubles.
* The converted AutoHotkey code will most likely be different for 32bit and 64bit (most notably: differences in structure offsets and variant handling in DLL-calls), so be sure to use the correct version and take care when manually merging the code for different bit versions.

Occasionally the type libraries make reference to interfaces or structures which are not included in the library. These can usually be found in other type libraries. Many basic structures and interfaces are defined in the type library "mscorlib.dll".


## Known issues ##

- some memory leaks

## Related ##

ImportTypeLib by maul.esel (https://github.com/maul-esel/ImportTypeLib) wraps type libraries directly at run-time. It requires a slightly more complex syntax in use and seems to have issues in 64bit.


## Example code AHK v1: ##

```ahk
; to use this example code first convert the type library UIAutomationClient with TypeLib2AHK

#Include UIAutomationClient_1_0_64bit.ahk  ; comment out as necessary
;~ #Include UIAutomationClient_1_0_32bit.ahk  ; uncomment as necessary

; Instantiate the CoClass; retrieves wrapper for IUIAutomation interface
UIA:=CUIAutomation()

; definition to set the correct variant type for CreatePropertyCondition below
Vt_Bool:=11

; Call an interface function; retrieves wrapper of the root IUIAutomationElement in the active window
CurrentWinRootElem := UIA.ElementFromHandle(WinExist("A"))

t:= "Root element:`n"
; gather information about the element
t.= ElemInfo(CurrentWinRootElem)
t.= "`n"

; The function CreateAndConditionFromArray expects a SAFEARRAY. The SAFEARRAY can be passed directly 
; or as shown here as an AHK object which is then converted by the wrapper
Conditions:=Object()
; only collect elements which are not offscreen
Conditions.Insert(UIA.CreatePropertyCondition(UIAutomationClientTLConst.UIA_IsOffscreenPropertyId, False, Vt_Bool))
; only collect elements which are controls
Conditions.Insert(UIA.ControlViewCondition)
; combine the above conditions with and
Condition := UIA.CreateAndConditionFromArray(Conditions)

; retrieve all decendants of the root element which meet the conditions
Descendants:=CurrentWinRootElem.FindAll(UIAutomationClientTLConst.TreeScope_Descendants, Condition)

; retrieve how many descendants were found
t.="Number of descendants: " Descendants.Length "`n`n"

; gather information about the first descendant element
t.= "First descendant element:`n"
; gather information about the first descendant element
t.=ElemInfo(Descendants.GetElement(1))

; display the result
msgbox, % t
return

ElemInfo(elem)
{
	; retrieve various properties of the IUIAutomationElement
	t := "Type: " UIAutomationClientTLConst.UIA_ControlTypeIds(elem.CurrentControlType) "`n"	
	t .= "Name:`t " elem.CurrentName "`n"
	; CurrentBoundingRectangle returns a tagRECT structure which is defined in the type library and also wrapped into an AHK object
	rect := elem.CurrentBoundingRectangle
	t .= "Location:`t Left: " rect.Left " Top: " rect.Top " Right: " rect.Right " Bottom: " rect.Bottom "`n"
	; GetClickablePoint expects a tagPOINT structure as parameter 
	; the structure can either be passed as an instance of the tagPOINT-wrapper as defined in the type library
	point:=new tagPOINT()
	out:=elem.GetClickablePoint(point)
	t.="Clickable point (using tagPoint structure): x: " point.x " y: " point.y " Has clickable point: " out "`n"
	; or it can be passed as a buffer of the right size
	VarSetCapacity(p,tagPOINT.__SizeOf(),0)
	out:=elem.GetClickablePoint(p)
	point:=new tagPOINT(p)
	t.="Clickable point (using buffer): x: " point.x " y: " point.y " Has clickable point: " out "`n"
	return t
}
```

## Example code AHK v2: ##

```ahk
; to use this example code first convert the type library UIAutomationClient with TypeLib2AHK

#Include UIAutomationClient_1_0_64bit.ahk2  ; comment out as necessary
;~ #Include UIAutomationClient_1_0_32bit.ahk2  ; uncomment as necessary

; Instantiate the CoClass; retrieves wrapper for IUIAutomation interface
UIA:=CUIAutomation()
; definition to set the correct variant type for CreatePropertyCondition below
Vt_Bool:=11

; Call an interface function; retrieves wrapper of the root IUIAutomationElement in the active window
CurrentWinRootElem := UIA.ElementFromHandle(WinExist("A"))

t:= "Root element:`n"
; gather information about the element
t.= ElemInfo(CurrentWinRootElem)
t.= "`n"

; The function CreateAndConditionFromArray expects a SAFEARRAY. The SAFEARRAY can be passed directly 
; or as shown here as an AHK object which is then converted by the wrapper
Conditions:=Object()
; only collect elements which are not offscreen
Conditions.Push(UIA.CreatePropertyCondition(UIAutomationClientTLConst.UIA_IsOffscreenPropertyId, False, Vt_Bool))
; only collect elements which are controls
Conditions.Push(UIA.ControlViewCondition)
; combine the above conditions with and
Condition := UIA.CreateAndConditionFromArray(Conditions)

; retrieve all decendants of the root element which meet the conditions
Descendants:=CurrentWinRootElem.FindAll(UIAutomationClientTLConst.TreeScope_Descendants, Condition)

; retrieve how many descendants were found
t.="Number of descendants: " Descendants.Length "`n`n"

; gather information about the first descendant element
t.= "First descendant element:`n"
; gather information about the first descendant element
t.=ElemInfo(Descendants.GetElement(1))

; display the result
msgbox t
return

ElemInfo(elem)
{
	; retrieve various properties of the IUIAutomationElement
	t := "Type: " UIAutomationClientTLConst.UIA_ControlTypeIds(elem.CurrentControlType) "`n"	
	t .= "Name:`t " elem.CurrentName "`n"
	; CurrentBoundingRectangle returns a tagRECT structure which is defined in the type library and also wrapped into an AHK object
	rect := elem.CurrentBoundingRectangle
	t .= "Location:`t Left: " rect.Left " Top: " rect.Top " Right: " rect.Right " Bottom: " rect.Bottom "`n"
	; GetClickablePoint expects a tagPOINT structure as parameter 
	; the structure can either be passed as an instance of the tagPOINT-wrapper as defined in the type library
	point:=new tagPOINT()
	out:=elem.GetClickablePoint(point)
	t.="Clickable point (using tagPoint structure): x: " point.x " y: " point.y " Has clickable point: " out "`n"
	; or it can be passed as a buffer of the right size
	VarSetCapacity(p,tagPOINT.__SizeOf(),0)
	out:=elem.GetClickablePoint(p)
	point:=new tagPOINT(p)
	t.="Clickable point (using buffer): x: " point.x " y: " point.y " Has clickable point: " out "`n"
	return t
}
```

