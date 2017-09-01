/*

  Version: MPL 2.0/GPL 3.0/LGPL 3.0

  The contents of this file are subject to the Mozilla Public License Version
  2.0 (the "License"); you may not use this file except in compliance with
  the License. You may obtain a copy of the License at

  http://www.mozilla.org/MPL/

  Software distributed under the License is distributed on an "AS IS" basis,
  WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
  for the specific language governing rights and limitations under the License.

  The Initial Developer of the Original Code is
  Elgin <Elgin_1@zoho.eu>.
  Portions created by the Initial Developer are Copyright (C) 2010-2017
  the Initial Developer. All Rights Reserved.

  Contributor(s):

  Alternatively, the contents of this file may be used under the terms of
  either the GNU General Public License Version 3 or later (the "GPL"), or
  the GNU Lesser General Public License Version 3.0 or later (the "LGPL"),
  in which case the provisions of the GPL or the LGPL are applicable instead
  of those above. If you wish to allow use of your version of this file only
  under the terms of either the GPL or the LGPL, and not to allow others to
  use your version of this file under the terms of the MPL, indicate your
  decision by deleting the provisions above and replace them with the notice
  and other provisions required by the GPL or the LGPL. If you do not delete
  the provisions above, a recipient may use your version of this file under
  the terms of any one of the MPL, the GPL or the LGPL.

*/

; ==============================================================================
; ==============================================================================
; known problems
; ==============================================================================
; ==============================================================================

; parameters of referenced type are not resolved if they point to another 
; referenced type -> probably never happens

; ==============================================================================
; ==============================================================================
; todo
; ==============================================================================
; ==============================================================================

; create code for referenced types not covered by the type library?

; ==============================================================================
; ==============================================================================
; Code
; ==============================================================================
; ==============================================================================

#SingleInstance Force
#Persistent
#Warn All, MsgBox
"".base.__Get := "".base.__Set := "".base.__Call := Func("Default__Warn")

; ==============================================================================
; Example: run to show the type information of a single COM object
;
;~ MainApp:=new TypeLib2AHKMain(ComObjCreate("InternetExplorer.Application"))
; ==============================================================================

; default: run to show a list of installed type libraries
MainApp:=new TypeLib2AHKMain()
return

;@Ahk2Exe-SetName TypeLib2AHK
;@Ahk2Exe-SetDescription TypeLib2AHK
;@Ahk2Exe-SetCopyright Elgin
;@Ahk2Exe-SetOrigFilename TypeLib2AHK.exe
;@Ahk2Exe-SetVersion 0.9.0.0

#Include .\Libraries
#Include TypeLibHelperFunctions.ahk
#Include TypeLibInterfaces.ahk
#Include ApplicationFramework.ahk
#Include VariousFunctions.ahk


; ==============================================================================
; ==============================================================================
; class TypeLib2AHK
;   __New(ApplicationName="NoName", DefaultUILanguage="English")
; ==============================================================================
; ==============================================================================


class TypeLib2AHKMain extends ApplicationFramework
{
	AppName:="TypeLib2AHK"
	Ver:="0.9.0.0"
	AuthorName:="Elgin"
	DefUILang:="English"
	
	BaseCOMObject:=0
	BaseCOMIndexInTypeLib:=0
	BaseCOMTypeInfo:=0
	BaseFileName:=""
	ActiveComObjects:=Object()
	
	__New(InObj=0)	; set InObj to a COM object to retrieve its type library information
	{
		base.__New(this.AppName, this.Ver, this.DefUILang)
		this.Settings.CreateIni("OutPutFolder",  A_MyDocuments)
		this.Settings.CreateIni("LoadFromFolder",  A_MyDocuments)
		this.Run(InObj)		
	}

	CleanUp(ExitReason, ExitCode)
	{
		base.CleanUp(ExitReason, ExitCode)
	}
	
	LaunchSelectTypeLibraryDialog()
	{
		this.STLGui:=this.ManagedGuis.NewGUI("TL2AHKSTL", this.Resources.SelectTypeLibraryDialogTitle, "+Resize +MaximizeBox +MinimizeBox")
		this.STLGui.OnClose:=ObjBindMethod(this,"STLClose")
		this.STLGui.OnEscape:=ObjBindMethod(this,"STLClose")

		this.STLLV:=this.STLGui.Add("ListView", "RegTypeLibs", this.Resources.SelectTypeLibraryDialogLVColumns, "r20 w700", ObjBindMethod(this,"STLLVEvent"))
		Loop, Reg, HKEY_CLASSES_ROOT\TypeLib, K
		{
			IID:=A_LoopRegName
			Loop, Reg, HKEY_CLASSES_ROOT\TypeLib\%IID%, K
			{
				RegRead, LibName, HKEY_CLASSES_ROOT\TypeLib\%IID%\%A_LoopRegName%
				If (LibName="")
				{
					VarSetCapacity(mem, 16, 00)
					hr:=DllCall("Ole32\CLSIDFromString", "Str", IID, "Ptr", &mem)
					RegExMatch(A_LoopRegName, "^(?P<Major>\d+)\.(?P<Minor>\d+)$", ver)
					lib:=0
					hr := DllCall("OleAut32\LoadRegTypeLib", "Ptr", &mem, "UShort", verMajor, "UShort", verMinor, "UInt", 0, "Ptr*", lib, "Int")
					If ((hr="" or hr>=0) and lib)
					{
						TypeLib:=new ITypeLib(lib)
						Doc:=TypeLib.GetDocumentation(-1)
						LibName:=Doc.Name
						If (LibName="")
							Libname:=Doc.DocString
						If (LibName="")
							LibName:=this.Resources.SelectTypeLibraryDialogNoNameStr		
					}
					else
						LibName:=""
				}
				If (LibName<>"")	; no need to add a library that can't be opened anyway
					this.STLLV.Add("",LibName,A_LoopRegName,IID)
			}
		}
		this.STLLV.ModifyCol(1, "Sort")
		pctinfo:=0
		out:=0
		cnt:=1
		for name, obj in GetActiveObjects()
		{			
			base:=ComObjValue(Obj)
			If (DllCall(NumGet(NumGet(base+0,"Ptr"), 03*A_PtrSize, "Ptr"), "Ptr", base, "UInt*", pctinfo, "Int")=0 and pctinfo>0 and DllCall(NumGet(NumGet(base+0,"Ptr"), 04*A_PtrSize, "Ptr"), "Ptr", base, "UInt", 0, "UInt", 0, "Ptr*", out, "Int")=0)
			{
				BaseTypeInfo:=new ITypeInfo(out)
				idoc:=BaseTypeInfo.GetDocumentation(-1)
				TypeLib:=BaseTypeInfo.GetContainingTypeLib(pindex)
				tdoc:=TypeLib.GetDocumentation(-1) ; MEMBERID_NIL
				tattr:=TypeLib.GetLibAttr()

				this.STLLV.Insert(cnt, "",Format(this.Resources.SelectTypeLibraryDialogRunningStr, idoc.Name, pindex, tdoc.Name, tdoc.DocString), tattr.wMajorVerNum "." tattr.wMinorVerNum, tattr.guid,cnt)
				TypeLib.ReleaseTLibAttr(tattr.__Ptr)
				this.ActiveComObjects[cnt]:=Object()
				this.ActiveComObjects[cnt].TypeLib:=TypeLib
				this.ActiveComObjects[cnt].Index:=pindex
				cnt++
			}
		}
		this.Settings.CreateIni("OutPutV2",  0)
		this.STLLV.ModifyCol(1, "AutoHdr")
		this.STLLV.ModifyCol(2, "AutoHdr")
		this.STLLV.ModifyCol(3, "AutoHdr")
		this.STLGui.Add("Button", "Convert", this.Resources.SelectTypeLibraryDialogButtonConvert, "", ObjBindMethod(this,"STLConvertButEvent"))
		this.STLGui.Add("Button", "View", this.Resources.SelectTypeLibraryDialogButtonView, "", ObjBindMethod(this,"STLSelectButEvent"))
		this.STLGui.Add("Button", "LoadFromFile", this.Resources.SelectTypeLibraryDialogButtonLoad, "x+m", ObjBindMethod(this,"STLLoadFromFileButEvent"))
		this.STLGui.Add("Button", "ConvertAll", this.Resources.SelectTypeLibraryDialogButtonConvertAll, "", ObjBindMethod(this,"STLConvertAllButEvent"))
		this.STLGui.Add("CheckBox", "CreateV2", this.Resources.SelectTypeLibraryDialogCheckV2, "Checked" this.Settings.OutPutV2, ObjBindMethod(this,"STLConvertCheckV2"), "CreateV2")
		this.STLGui.AddSizingInfo("XMin", ["Convert", "View", "LoadFromFile", "ConvertAll", "CreateV2"], 1, "D", 0)
		this.STLGui.AddSizingInfo("YMin", "RegTypeLibs", 1, 200, 0, "*", 0, "M", 0, ["Convert", "View", "LoadFromFile", "ConvertAll", "CreateV2"], 0, "C", 0)
		this.STLGui.AddSizingInfo("X", "RegTypeLibs", 1, 100, 0)
		this.STLGui.Show()
	}
	
	Run(InObj=0)
	{
		If (!InObj)
		{
			this.LaunchSelectTypeLibraryDialog()
		}
		else
		{
			this.BaseCOMObject:=ComObjValue(InObj)
			pctinfo:=0
			out:=0
			If (DllCall(NumGet(NumGet(this.BaseCOMObject+0,"Ptr"), 03*A_PtrSize, "Ptr"), "Ptr", this.BaseCOMObject, "UInt*", pctinfo, "Int")=0 and pctinfo>0 and DllCall(NumGet(NumGet(this.BaseCOMObject+0,"Ptr"), 04*A_PtrSize, "Ptr"), "Ptr", this.BaseCOMObject, "UInt", 0, "UInt", 0, "Ptr*", out, "Int")=0)
			{
				this.BaseCOMTypeInfo:=new ITypeInfo(out)
				this.TypeLib:=this.BaseCOMTypeInfo.GetContainingTypeLib(pindex)
				this.BaseCOMIndexInTypeLib:=pindex
				this.ShowInfo(pindex)
			}
			else
			{
				this.BaseComObject:=0
				msgbox, % this.Resources.LoadCOMFailed
				this.LaunchSelectTypeLibraryDialog()
			}
		}
	}
	
	ShowInfo(Index="")
	{
		;~ this.TLShow:=TypeLibToVerboseObj(this.TypeLib, Index) ; show complete information of the type library; very slow and hard to look through...
		this.TLShow:=TypeLibToCondensedObj(this.TypeLib, Index) ; show basic, formatted information of the type library
		If (this.TLShow)
		{
			this.TLIGui:=this.ManagedGuis.NewGUI("TL2AHKSTL", "TypeLib2AHK - View Type Library", "+Resize +MaximizeBox +MinimizeBox")
			this.TLIGui.OnClose:=ObjBindMethod(this,"TLIClose")
			this.TLIGui.OnEscape:=ObjBindMethod(this,"TLIClose")
			this.TLITV:=this.TLIGui.Add("TreeView", "TypeLibsView", "", "r20 w700", ObjBindMethod(this,"TLITVEvent"))
			this.TLIGui.Add("Button", "Convert", this.Resources.SelectTypeLibraryDialogButtonConvert, "", ObjBindMethod(this,"TLIConvertButEvent"))
			this.TLIGui.Add("Button", "Cancel", this.Resources.SelectTypeLibraryDialogButtonCancel, "x+m", ObjBindMethod(this,"TLIClose"))
			this.TLIGui.AddSizingInfo("XMin", ["Convert","Cancel"], 1, "D", 0)
			this.TLIGui.AddSizingInfo("YMin", "TypeLibsView", 1, 200, 0, "*", 0, "M", 0, ["Convert","Cancel"], 0, "C", 0)
			this.TLIGui.AddSizingInfo("X", "TypeLibsView", 1, 100, 0)
			this.TLIGui.Show()			
			this.TLITV.Add(this.Resources.TreeUpdateMsg)
			sleep, 10
			this.TLITV.Control("-Redraw")
			this.TLITV.Delete()
			ObjToTreeView(this.TLShow, ObjBindMethod(this,"TLITVFormatFunc"))
			this.TLITV.Control("+Redraw")
		}
	}
	
	STLClose()
	{
		ExitApp
	}
	
	STLConvertCheckV2(CtrlHwnd="", GuiEvent="", EventInfo="", ErrorLvl="")
	{
		this.Settings.OutPutV2:=this.STLGui.Submit().CreateV2
	}
	
	STLLoadFromFileButEvent(CtrlHwnd="", GuiEvent="", EventInfo="", ErrorLvl="")
	{
		FileSelectFile, FileName, , % this.Settings.LoadFromFolder, % this.Resources.LoadFileSelect, % this.Resources.LoadFileFiles
		If (FileName<>"")
		{
			ResourceNumber:=1
			matchcnt:=RegExMatch(FileName, "O)_([0-9]+)$", match)
			if (matchcnt>0)
			{
				ResourceNumber:=match.Value(1)
				FileName:=Substr(FileName,1,match.Pos(1)-2)				
			}
			SplitPath, FileName, , OutDir
			this.Settings.LoadFromFolder:=OutDir
			lib:=0
			hr := DllCall("OleAut32\LoadTypeLib", "Str", FileName "\" ResourceNumber, "Ptr*", lib, "Int") 
			If (!hr)
			{
				this.TypeLib:=new ITypeLib(lib)
				this.BaseFileName:=FileName
				this.ShowInfo()
			}
			else
			{
				If (hr=-2147319779)
					msgbox, % this.Resources.LoadFileFailed FileName " : Library not registered."
				else
					msgbox, % this.Resources.LoadFileFailed FileName " : 0x" Format("{:x}", hr & 0xFFFFFFFF) " (" hr ")"
			}
			
		}		
	}
	
	STLLVEvent(CtrlHwnd="", GuiEvent="", EventInfo="", ErrorLvl="")
	{
		If (GuiEvent="DOUBLECLICK")
		{
			this.STLSelectButEvent()
		}
	}
	
	STLConvertButEvent(CtrlHwnd="", GuiEvent="", EventInfo="", ErrorLvl="")
	{
		this.STLGui.SelectViewControl("RegTypeLibs")
		row:=this.STLLV.GetNext(0,"Focused")
		if (this.ActiveComObjects.Haskey(row))
		{
			this.TypeLib:=this.ActiveComObjects[row].TypeLib
			this.TLIConvertButEvent()
		}		
		else
		{
			this.STLLV.GetText(version,row,2)
			this.STLLV.GetText(GUID,row,3)
			VarSetCapacity(mem, 16, 00)
			hr:=DllCall("Ole32\CLSIDFromString", "Str", GUID, "Ptr", &mem)
			RegExMatch(version, "^(?P<Major>\d+)\.(?P<Minor>\d+)$", ver)
			lib:=0
			hr := DllCall("OleAut32\LoadRegTypeLib", "Ptr", &mem, "UShort", verMajor, "UShort", verMinor, "UInt", 0, "Ptr*", lib, "Int")
			If ((hr="" or hr>=0) and lib)
			{
				this.TypeLib:=new ITypeLib(lib)
				this.TLIConvertButEvent()
			}
			else
			{
				If (hr=-2147319779)
					msgbox, % this.Resources.LoadFileFailed " Library not registered."
				else
					msgbox, % this.Resources.LoadFileFailed " 0x" Format("{:x}", hr & 0xFFFFFFFF) " (" hr ")"
			}
		}
	}
	
	STLConvertAllButEvent(CtrlHwnd="", GuiEvent="", EventInfo="", ErrorLvl="")	; existing files will not be overwritten
	{
		FileSelectFolder, FolderName, % "*" this.Settings.OutPutFolder
		If (ErrorLevel=1)
			return
		this.Settings.OutPutFolder:=RegExReplace(FolderName, "\\$")  
		If (FileExist(this.Settings.OutPutFolder)<>"D")
			FileCreateDir, % this.Settings.OutPutFolder
		Loop, Reg, HKEY_CLASSES_ROOT\TypeLib, K
		{
			IID:=A_LoopRegName
			Loop, Reg, HKEY_CLASSES_ROOT\TypeLib\%IID%, K
			{
				VarSetCapacity(mem, 16, 00)
				hr:=DllCall("Ole32\CLSIDFromString", "Str", IID, "Ptr", &mem)
				RegExMatch(A_LoopRegName, "^(?P<Major>\d+)\.(?P<Minor>\d+)$", ver)
				lib:=0
				hr := DllCall("OleAut32\LoadRegTypeLib", "Ptr", &mem, "UShort", verMajor, "UShort", verMinor, "UInt", 0, "Ptr*", lib, "Int")
				If ((hr="" or hr>=0) and lib)
				{
					this.TypeLib:=new ITypeLib(lib)
					this.ConvertTL()
				}
			}
		}
	}
	
	STLSelectButEvent(CtrlHwnd="", GuiEvent="", EventInfo="", ErrorLvl="")
	{
		this.STLGui.SelectViewControl("RegTypeLibs")
		row:=this.STLLV.GetNext(0,"Focused")
		if (this.ActiveComObjects.Haskey(row))
		{
			this.TypeLib:=this.ActiveComObjects[row].TypeLib
			this.ShowInfo(this.ActiveComObjects[row].Index)
		}		
		else
		{
			this.STLLV.GetText(version,row,2)
			this.STLLV.GetText(GUID,row,3)
			VarSetCapacity(mem, 16, 00)
			hr:=DllCall("Ole32\CLSIDFromString", "Str", GUID, "Ptr", &mem)
			RegExMatch(version, "^(?P<Major>\d+)\.(?P<Minor>\d+)$", ver)
			lib:=0
			hr := DllCall("OleAut32\LoadRegTypeLib", "Ptr", &mem, "UShort", verMajor, "UShort", verMinor, "UInt", 0, "Ptr*", lib, "Int") ; error handling is done below
			If ((hr="" or hr>=0) and lib)
			{
				this.TypeLib:=new ITypeLib(lib)
				this.ShowInfo()
			}
			else
			{
				If (hr=-2147319779)
					msgbox, % this.Resources.LoadFileFailed " Library not registered."
				else
					msgbox, % this.Resources.LoadFileFailed " 0x" Format("{:x}", hr & 0xFFFFFFFF) " (" hr ")"
			}
		}
	}
	
	TLITVFormatFunc(Depth, Index, Value)
	{
		Out:=Object()
		If (IsObject(Value))
		{
			Out.Text:=Index
			Out.RecurseInto:=1
		}
		else
		{
			Out.Text:=Index ": " Value
			Out.RecurseInto:=0
		}
		If (Depth<2)
			Out.Options:="Expand"
		else
			Out.Options:=""
		Out.Add:=1
		Out.Continue:=1
		return Out
	}
	
	TLIClose()
	{
		this.TLIGui.Hide()
	}
	
	TLIConvertButEvent(CtrlHwnd="", GuiEvent="", EventInfo="", ErrorLvl="")
	{
		this.TLInfo:=TypeLibToHeadingsObj(this.TypeLib)
		bit:=A_PtrSize*8
		FileSelectFile, FileName, S 18, % this.Settings.OutPutFolder "\" this.TLInfo._TypeLibraryInfo.Name "_" this.TLInfo._TypeLibraryInfo.wMajorVerNum "_" this.TLInfo._TypeLibraryInfo.wMinorVerNum "_" bit "bit.ahk" (this.Settings.OutPutV2 ? "2":"")
		If (FileName<>"")
		{
			SplitPath, FileName, OutFileName, OutDir
			this.Settings.OutPutFolder:=OutDir
			File:=FileOpen(FileName, "w")
			File.Write(this.MakeHeader(OutFileName))
			If (IsObject(this.TLInfo.COCLASS)) 
				File.Write(this.MakeCoClass())
			If (IsObject(this.TLInfo.ALIAS)) 
				File.Write(this.MakeGeneric("ALIAS"))
			If (IsObject(this.TLInfo.ENUM)) 
				File.Write(this.MakeConst())
			If (IsObject(this.TLInfo.RECORD)) 
				File.Write(this.MakeRecord("Record"))
			If (IsObject(this.TLInfo.UNION)) 
				File.Write(this.MakeRecord("Union"))
			If (IsObject(this.TLInfo.INTERFACE)) 
				File.Write(this.MakeInterface())
			If (IsObject(this.TLInfo.DISPATCH)) 
				File.Write(this.MakeGeneric("DISPATCH"))
			File.Close()
			this.TLIGui.Close()
		}
	}
	
	ConvertTL()
	{
		this.TLInfo:=TypeLibToHeadingsObj(this.TypeLib)
		bit:=A_PtrSize*8
		FileName:=this.Settings.OutPutFolder "\" this.TLInfo._TypeLibraryInfo.Name "_" this.TLInfo._TypeLibraryInfo.wMajorVerNum "_" this.TLInfo._TypeLibraryInfo.wMinorVerNum "_" bit "bit.ahk" (this.Settings.OutPutV2 ? "2":"")
		If (FileExist(FileName))
			return
		If (FileName<>"")
		{
			SplitPath, FileName, OutFileName, OutDir
			this.Settings.OutPutFolder:=OutDir
			File:=FileOpen(FileName, "w")
			File.Write(this.MakeHeader(OutFileName))
			If (IsObject(this.TLInfo.COCLASS)) 
				File.Write(this.MakeCoClass())
			If (IsObject(this.TLInfo.ALIAS)) 
				File.Write(this.MakeGeneric("ALIAS"))
			If (IsObject(this.TLInfo.ENUM)) 
				File.Write(this.MakeConst())
			If (IsObject(this.TLInfo.RECORD)) 
				File.Write(this.MakeRecord("Record"))
			If (IsObject(this.TLInfo.UNION)) 
				File.Write(this.MakeRecord("Union"))
			If (IsObject(this.TLInfo.INTERFACE)) 
				File.Write(this.MakeInterface())
			If (IsObject(this.TLInfo.DISPATCH)) 
				File.Write(this.MakeGeneric("DISPATCH"))
			File.Close()
		}
	}
	
	MakeHeader(FileName)
	{
		t:=this.Resources.StartComment
		t.=Format(this.Resources.IntroComment1,FileName,this.TLInfo._TypeLibraryInfo.Name,this.TLInfo._TypeLibraryInfo.wMajorVerNum,this.TLInfo._TypeLibraryInfo.wMinorVerNum, this.TLInfo._TypeLibraryInfo.DocString,this.AppName, this.Ver, this.AuthorName)
		s:=""
		for, index, key in this.TLInfo._TypeLibraryInfo.LibFlagsNames
		  s.=", " SubStr(key, 9)
		t.=Format(this.Resources.IntroComment2, this.TLInfo._TypeLibraryInfo.guid, this.TLInfo._TypeLibraryInfo.lcid,this.TLInfo._TypeLibraryInfo.HelpFile, this.TLInfo._TypeLibraryInfo.HelpContext, SubStr(SYSKIND(this.TLInfo._TypeLibraryInfo.SysKind),5), SubStr(s, 3))
		t.=this.Resources.EndComment
		return t
	}
	
	MakeCoClass()
	{
		t:=this.Resources.StartComment this.Resources.StartComment this.Resources.IntroCoClass this.Resources.StartComment this.Resources.EndComment
		for index, obj in this.TLInfo["COCLASS"]
		{
			Progress, % (Index)/this.TLInfo["COCLASS"].MaxIndex()*100, % Index this.Resources.Of this.TLInfo["COCLASS"].MaxIndex(), % this.Resources.ProgressWindowText "CoClass", % this.Resources.ProgressWindowCaption ": " this.TLInfo._TypeLibraryInfo.Name
			TI:=this.TypeLib.GetTypeInfo(obj.Index)
			t.=this.Resources.StartComment "; " TI.GetDocumentation(-1).Name	"`r`n" this.Resources.EndComment
			Attr:=TI.GetTypeAttr()
			t.=this.Resources.StartBlockComment this.Resources.GUID Attr.guid "`r`n"
			DefName:=""
			If (Attr.cImplTypes)
			{
				If (Attr.typekind=5)	; TKIND_COCLASS
				{
					Loop, % Attr.cImplTypes
					{
						ImplT:=TI.GetRefTypeOfImplType(A_Index-1)
						Name:=ImplT.GetDocumentation(-1).Name
						t.=this.Resources.Implements Name "; "
						ImplAttr:=ImplT.GetTypeAttr()
						t.=this.Resources.GUID ImplAttr.guid "; "
						t.=this.Resources.Flags 
						for index, flag in IMPLTYPEFLAGS(TI.GetImplTypeFlags(A_Index-1))
						{
							If (index>1)
								t.=", "
							t.=flag
						}
						t.="`r`n"
						If (TI.GetImplTypeFlags(A_Index-1) & 0x1) ; IMPLTYPEFLAG_FDEFAULT
						{
							DefName:=Name
							DefGUID:=ImplAttr.guid
							DefTypeKind:=ImplAttr.typekind
						}
						ImplAttr.ReleaseTypeAttr(ImplAttr.__Ptr)
					}
				}
			}
			TI.ReleaseTypeAttr(Attr.__Ptr)
			t.=this.Resources.EndBlockComment
			If (DefName<>"")
			{
				t.=TI.GetDocumentation(-1).Name "()`r`n{`r`n	try`r`n	{`r`n		If (impl:=ComObjCreate(""" Attr.guid ""","""DefGUID """))`r`n"
				If (DefTypeKind=3) ; Interface
					t.="			return new " DefName "(impl)`r`n"
				else
					t.="			return impl`r`n"
				t.="		throw """ DefName this.Resources.CodeComNotInitialized
				t.="	}`r`n	catch e`r`n"
				If (this.Settings.OutPutV2)
					t.="		MsgBox, 262160, " DefName " Error, IsObject(e)?""" DefName this.Resources.CodeComNotRegistered ":e.Message`r`n"
				else
					t.="		MsgBox, 262160, " DefName " Error, % IsObject(e)?""" DefName this.Resources.CodeComNotRegistered ":e.Message`r`n"
				t.="}`r`n`r`n"						
			}					
		}
		Progress, Off
		return t
	}
	
	MakeConst()
	{
		t:=this.Resources.StartComment
		t.=this.Resources.StartComment
		t.=this.Resources.IntroENUM1
		t.=this.Resources.StartComment
		t.=this.Resources.EndComment
		t.="class " this.TLInfo._TypeLibraryInfo.Name "TLConst`r`n{`r`n"
		for index, obj in this.TLInfo.ENUM
		{
			Progress, % (Index)/this.TLInfo.ENUM.MaxIndex()*100, % Index " of " this.TLInfo.ENUM.MaxIndex(), % this.Resources.ProgressWindowText "Enum", % this.Resources.ProgressWindowCaption ": " this.TLInfo._TypeLibraryInfo.Name
			TI:=this.TypeLib.GetTypeInfo(obj.Index)
			t.="`t" this.Resources.StartComment
			t.="`t" Format(this.Resources.IntroENUM2, TI.GetDocumentation(-1).Name)
			t.="`t" this.Resources.EndComment
			Attr:=TI.GetTypeAttr()
			If (Attr.cVars)
			{
				vcount:=1
				tRev:="`t" TI.GetDocumentation(-1).Name "(Value)`r`n	{`r`n		static v1:={"
				Loop, % Attr.cVars
				{
					If (Mod(A_Index,126)=0) ; need several variables to store the data due to the limit of 512 operators and operands per expression
					{
						vcount++
						tRev.="}`r`n		static v" vcount ":={"
					}
					VarDescVar:=TI.GetVarDesc(A_Index-1)
					Doc:=TI.GetDocumentation(VarDescVar.memid)
					If (VarDescVar.varkind=2) ; Const
					{
						
						t.="`tstatic " Doc.Name " := "  Format("0x{1:X}", VarDescVar.oInst) ; "`r`n"
						If (Doc.DocString<>"")
							 t.="  `; " Doc.DocString 
						t.="`r`n"
						tRev.=(A_Index>1 and Mod(A_Index,126)<>0) ? ", " : "" 
						tRev.=Format("0x{1:X}", VarDescVar.oInst) ":""" Doc.Name """"
					}
					TI.ReleaseVarDesc(VarDescVar.__Ptr)
				}
				t.="`r`n" tRev "}`r`n"
				Loop, %vcount%
				{
					If (A_Index>1)
						t.="		else`r`n"
					t.="		If (v" A_Index "[Value])`r`n"
					t.="			return v" A_Index "[Value]`r`n"
				}
				t.="	}`r`n"
			}
			TI.ReleaseTypeAttr(Attr.__Ptr)
			t.="`r`n"
		}
		for index, obj in this.TLInfo.MODULE
		{
			Progress, % (Index)/this.TLInfo.MODULE.MaxIndex()*100, % Index " of " this.TLInfo.MODULE.MaxIndex(), % this.Resources.ProgressWindowText "Module", % this.Resources.ProgressWindowCaption ": " this.TLInfo._TypeLibraryInfo.Name
			TI:=this.TypeLib.GetTypeInfo(obj.Index)
			t.="`t" this.Resources.StartComment
			t.="`t" Format(this.Resources.IntroMODULE2, TI.GetDocumentation(-1).Name)
			t.="`t" this.Resources.EndComment
			Attr:=TI.GetTypeAttr()
			If (Attr.cFuncs)
			{
				t.=this.Resources.StartBlockComment
				Loop, % Attr.cFuncs
				{
					FuncIndex:=A_Index-1
					FuncDescVar:=TI.GetFuncDesc(FuncIndex)
					Doc:=TI.GetDocumentation(FuncDescVar.memid)
					Name:=Doc.Name
					If Name not in QueryInterface,AddRef,Release,GetTypeInfoCount,GetTypeInfo,GetIDsOfNames,Invoke
					{
						t.=this.Resources.VTablePositon ": " FuncDescVar.oVft//A_PtrSize ":`r`n"
						t.=INVOKEKIND(FuncDescVar.invkind) " "
						t.=SubStr(VARENUM(FuncDescVar.elemdescFunc.tdesc.vt), 4) " "
						t.=Name "("
						ParamNames:=TI.GetNames(FuncDescVar.memid,FuncDescVar.cParams+1)
						Loop, % FuncDescVar.cParams
						{
							If (A_Index>1)
								t.=", "
							param:=new ELEMDESC(FuncDescVar.lprgelemdescParam+(A_Index-1)*ELEMDESC.SizeOf())
							for index, key in PARAMFLAG(param.paramdesc.wParamFlags)
								t.="[" SubStr(key, 11) "] "
							vstr:=GetVarStr(param, TI)
							If (InStr(vstr,"VT_")=1)
								t.=Substr(vstr, 4) ": "
							else
								t.=vstr ": "
							t.=ParamNames[A_Index+1]
							If (param.paramdesc.wParamFlags & 0x20) ; Hasdefault
							{
								t.=" = " param.paramdesc.pPARAMDescEx.varDefaultValue
							}
						}
						t.=")`r`n"
						If (Doc.DocString<>"")
							t.=Doc.DocString "`r`n"
						t.="`r`n"
					}
					TI.ReleaseFuncDesc(FuncDescVar.__Ptr)
				}
				t.=this.Resources.StartBlockComment
			}			
			If (Attr.cVars)
			{
				vcount:=1
				tRev:="`t" TI.GetDocumentation(-1).Name "(Value)`r`n	{`r`n		static v1:={"
				Loop, % Attr.cVars
				{
					If (Mod(A_Index,126)=0) ; need several variables to store the data due to the limit of 512 operators and operands per expression
					{
						vcount++
						tRev.="}`r`n		static v" vcount ":={"
					}
					VarDescVar:=TI.GetVarDesc(A_Index-1)
					Doc:=TI.GetDocumentation(VarDescVar.memid)
					If (VarDescVar.varkind=2) ; Const
					{
						vs:=GetVarStr(VarDescVar.elemdescVar, TI)
						If (vs="Vt_Bstr")							
							t.="`tstatic " Doc.Name " := """  VarDescVar.oInst """  `; Type: " vs ; "`r`n"
						else
							t.="`tstatic " Doc.Name " := "  VarDescVar.oInst "  `; Type: " vs ; "`r`n"
						If (Doc.DocString<>"")
							 t.="  `; " Doc.DocString
						t.="`r`n"
						tRev.=(A_Index>1 and Mod(A_Index,126)<>0) ? ", " : "" 
						tRev.=Format("0x{1:X}", VarDescVar.oInst) ":""" Doc.Name """"
					}
					TI.ReleaseVarDesc(VarDescVar.__Ptr)
				}
				t.="`r`n" tRev "}`r`n"
				Loop, %vcount%
				{
					If (A_Index>1)
						t.="		else`r`n"
					t.="		If (v" A_Index "[Value])`r`n"
					t.="			return v" A_Index "[Value]`r`n"
				}
				t.="	}`r`n"
			}
			TI.ReleaseTypeAttr(Attr.__Ptr)
			t.="`r`n"
		}
		Progress, Off
		t.="}`r`n`r`n"
		return t
	}
	
	MakeInterface()
	{
		t:=this.Resources.StartComment this.Resources.StartComment this.Resources.IntroINTERFACE this.Resources.StartComment this.Resources.EndComment
		for index, obj in this.TLInfo.INTERFACE
		{
			Progress, % (Index)/this.TLInfo.INTERFACE.MaxIndex()*100, % Index this.Resources.Of this.TLInfo.INTERFACE.MaxIndex(), % this.Resources.ProgressWindowText "Interface", % this.Resources.ProgressWindowCaption ": " this.TLInfo._TypeLibraryInfo.Name
			TI:=this.TypeLib.GetTypeInfo(obj.Index)
			t.=this.Resources.StartComment
			t.="; " TI.GetDocumentation(-1).Name	"`r`n"
			Attr:=TI.GetTypeAttr()
			t.="; " this.Resources.GUID Attr.guid "`r`n"
			t.=this.Resources.EndComment
			t.="`r`n"			
			extends:=""
			If (Attr.cImplTypes)
			{
				If (Attr.typekind=3 or Attr.typekind=4 or Attr.typekind=5)	; TKIND_INTERFACE, TKIND_DISPATCH, TKIND_COCLASS
				{
					Loop, % Attr.cImplTypes
					{
						Name:=TI.GetRefTypeOfImplType(A_Index-1).GetDocumentation(-1).Name
						If (Name<>"IUnknown")
							extends.=Name
					}
				}
				else
				If (Attr.typekind=6)	; TKIND_ALIAS
				{
					t.=this.Resources.StartComment
					t.=this.Resources.Alias GetVarStrFromTD(Attr.tdescAlias, TI)
					t.=this.Resources.EndComment
				}
				else
				{
					Loop, % Attr.cImplTypes
					{
						Name:=TI.GetRefTypeOfImplType(A_Index-1).GetDocumentation(-1).Name
						t.=this.Resources.StartComment
						t.=this.Resources.Implements Name "`r`n"
						t.=this.Resources.EndComment
					}
				}
			}			
			t.="class " TI.GetDocumentation(-1).Name
			if (extends)
				t.=" extends " extends "`r`n{" 
			else
			{
				t.="`r`n{ `r`n	" this.Resources.CodeGeneric
				t.="	static __IID := """ Attr.guid """`r`n" "`r`n"
				If (this.Settings.OutPutV2)
					t.="	__New(p:="""", flag:=1)`r`n	{`r`n		this.__Type:=""" TI.GetDocumentation(-1).Name """`r`n		this.__Value:=p`r`n		this.__Flag:=flag`r`n	}`r`n`r`n	__Delete()`r`n	{`r`n		this.__Flag? ObjRelease(this.__Value):0`r`n	}`r`n`r`n	__Vt(n)`r`n	{`r`n		return NumGet(NumGet(this.__Value+0, ""Ptr"")+n*A_PtrSize,""Ptr"")`r`n	}`r`n"
				else
					t.="	__New(p="""", flag=1)`r`n	{`r`n		this.__Type:=""" TI.GetDocumentation(-1).Name """`r`n		this.__Value:=p`r`n		this.__Flag:=flag`r`n	}`r`n`r`n	__Delete()`r`n	{`r`n		this.__Flag? ObjRelease(this.__Value):0`r`n	}`r`n`r`n	__Vt(n)`r`n	{`r`n		return NumGet(NumGet(this.__Value+0, ""Ptr"")+n*A_PtrSize,""Ptr"")`r`n	}`r`n"
					;~ t.="	__New(p="""", flag=1)`r`n	{`r`n		ObjInsert(this, ""__Type"", """ TI.GetDocumentation(-1).Name """)`r`n			,ObjInsert(this, ""__Value"", p)`r`n			,ObjInsert(this, ""__Flag"", flag)`r`n	}`r`n`r`n	__Delete()`r`n	{`r`n		this.__Flag? ObjRelease(this.__Value):0`r`n	}`r`n`r`n	__Vt(n)`r`n	{`r`n		return NumGet(NumGet(this.__Value+0, ""Ptr"")+n*A_PtrSize,""Ptr"")`r`n	}`r`n"
			}
			t.="`r`n"
			If (Attr.cVars)
			{
				t.="	" this.Resources.CodeInterfaceConstants
				Loop, % Attr.cVars
				{
					VarDescVar:=TI.GetVarDesc(A_Index-1)
					Doc:=TI.GetDocumentation(VarDescVar.memid)
					If (VarDescVar.varkind=2) ; Const
					{
						
						t.="`t" Doc.Name " := "  Format("0x{1:X}", VarDescVar.oInst) ; "`r`n"
						If (Doc.DocString<>"")
							 t.="  `; " Doc.DocString 
						t.="`r`n"
					}
					TI.ReleaseVarDesc(VarDescVar.__Ptr)
				}
			}
			If (Attr.cFuncs)
			{
				PropertyStore:=Object()
				t.="	" this.Resources.CodeInterfaceFunctions
				Loop, % Attr.cFuncs
				{
					FuncIndex:=A_Index-1
					FuncDescVar:=TI.GetFuncDesc(FuncIndex)
					Doc:=TI.GetDocumentation(FuncDescVar.memid)
					Name:=Doc.Name
					If (FuncDescVar.invkind=1)	; Is a function not a property
					{
						ReturnIsHRESULT:=0
						HasRetValParam:=0
						RetValTypeObj:=
						RetValName:=""
						t.="	" "; " this.Resources.VTablePositon " " FuncDescVar.oVft//A_PtrSize ": "
						t.=INVOKEKIND(FuncDescVar.invkind) " "
						result:=FuncDescVar.elemdescFunc.tdesc.vt
						If (result=25) ; VT_HRESULT
							ReturnIsHRESULT:=1
						t.=VARENUM(result) " "							
						t.=Name "("
						ParamNames:=TI.GetNames(FuncDescVar.memid,FuncDescVar.cParams+1)
						Loop, % FuncDescVar.cParams
						{
							If (A_Index>1)
								t.=", "
							param:=new ELEMDESC(FuncDescVar.lprgelemdescParam+(A_Index-1)*ELEMDESC.SizeOf())
							vstr:=GetVarStr(param, TI)
							If (param.paramdesc.wParamFlags & 0x8)
							{
								HasRetValParam:=1								
								RetValName:=ParamNames[A_Index+1]
								RetValTypeObj:=GetTypeObj(param, TI)
							}
							for index, key in PARAMFLAG(param.paramdesc.wParamFlags)
							{
								t.="[" SubStr(key, 11) "] "
							}
							If (InStr(vstr,"VT_")=1)
								t.=Substr(vstr, 4) ": "
							else
								t.=vstr ": "
							t.=ParamNames[A_Index+1]
							If (param.paramdesc.wParamFlags & 0x20) ; Hasdefault
							{
								t.=" = " param.paramdesc.pPARAMDescEx.varDefaultValue
							}
						}
						t.=")`r`n"
						If (Doc.DocString<>"")
							t.="	`; " Doc.DocString "`r`n"
						
						t.="	" Name "("
						Loop, % FuncDescVar.cParams
						{
							param:=new ELEMDESC(FuncDescVar.lprgelemdescParam+(A_Index-1)*ELEMDESC.SizeOf())
							If (ReturnIsHRESULT and HasRetValParam and (param.paramdesc.wParamFlags & 0x8)) ; skip PARAMFLAG_FRETVAL
								continue
							If (A_Index>1)
								t.=", "
							If (param.paramdesc.wParamFlags & 0x2)	; PARAMFLAG_FOUT
								t.="byref "
							t.=ParamNames[A_Index+1]
							If (param.paramdesc.wParamFlags & 0x20) ; Hasdefault
							{
								If (this.Settings.OutPutV2)
									t.=" := " param.paramdesc.pPARAMDescEx.varDefaultValue
								else
									t.=" = " param.paramdesc.pPARAMDescEx.varDefaultValue
							}
							else
							If (param.paramdesc.wParamFlags & 0x10) ; PARAMFLAG_FOPT
							{
								If (this.Settings.OutPutV2)
									t.=":=0"
								else
									t.="=0"
							}
							If (param.tdesc.vt=12) ; VT_Variant
								t.=", " ParamNames[A_Index+1] "VariantType"							
						}
						t.=")`r`n	{`r`n"							
						Loop, % FuncDescVar.cParams
						{
							param:=new ELEMDESC(FuncDescVar.lprgelemdescParam+(A_Index-1)*ELEMDESC.SizeOf())
							If (param.tdesc.vt=12) ; VT_Variant
							{
								If (A_PtrSize=8)
								{
									t.="		if (" ParamNames[A_Index+1] "VariantType!=12) `; " ParamNames[A_Index+1] " is not a variant`r`n"
									t.="		{`r`n"
									t.="			VarSetCapacity(ref" ParamNames[A_Index+1] ",8+2*A_PtrSize)`r`n"
									t.="			variant_ref := ComObject(0x400C, &ref" ParamNames[A_Index+1] ")`r`n"
									t.="			variant_ref[] := " ParamNames[A_Index+1] "`r`n"
									t.="			NumPut(" ParamNames[A_Index+1] "VariantType, ref" ParamNames[A_Index+1] ", 0, ""short"")`r`n"
									t.="		}`r`n"
									t.="		else`r`n"
									t.="			ref" ParamNames[A_Index+1] ":=" ParamNames[A_Index+1] "`r`n"
								}
								else
								{
									t.="		if (" ParamNames[A_Index+1] "VariantType=8)`r`n"
									t.="		" ParamNames[A_Index+1] ":=DllCall(""oleaut32\SysAllocString"", ""wstr"", ref" ParamNames[A_Index+1] ",""Ptr"")`r`n"
								}		
							}
						}		
						Loop, % FuncDescVar.cParams
						{
							param:=new ELEMDESC(FuncDescVar.lprgelemdescParam+(A_Index-1)*ELEMDESC.SizeOf())
							to:=GetTypeObj(param, TI)
							If ((to[1].Type=26 and to[2].Type=27) or (to[1].Type=27)) ; create code to handle safearrays
							{
								t.="		If (ComObjValue(" ParamNames[A_Index+1] ") & 0x2000)`r`n"
								t.="			ref" ParamNames[A_Index+1] ":=" ParamNames[A_Index+1] "`r`n"
								t.="		else`r`n"
								t.="		{`r`n"
								t.="			ref" ParamNames[A_Index+1] ":=ComObject(0x2003, DllCall(""oleaut32\SafeArrayCreateVector"", ""UInt"", 13, ""UInt"", 0, ""UInt"", " ParamNames[A_Index+1] ".MaxIndex()),1)`r`n"
								t.="			For ind, val in " ParamNames[A_Index+1] "`r`n"
								t.="				ref" ParamNames[A_Index+1] "[A_Index-1]:= val.__Value, ObjAddRef(val.__Value)`r`n"
								t.="		}`r`n"
							}
							If ((to[1].Type=26 and to[2].Type=29 and to[2].RefType=13) or (to[1].Type=29 and to[1].RefType=13)) ; create code to handle referenced types
							{
								If ((to[1].Type=26 and to[2].Type=29 and to[2].IsInterface=1) or (to[1].Type=29 and to[1].IsInterface=1))
								{
									; Check if the passed parameter is an object handled by this converted library and pass the referenced COM object
									t.="		If (IsObject(" ParamNames[A_Index+1] ") and (ComObjType(" ParamNames[A_Index+1] ")=""""))`r`n"
									t.="			ref" ParamNames[A_Index+1] ":=" ParamNames[A_Index+1] ".__Value`r`n"
									; If the parameter is a COM object, pass it through
									t.="		else`r`n"
									t.="			ref" ParamNames[A_Index+1] ":=" ParamNames[A_Index+1] "`r`n"
								}
								else
								{
									; Check if the passed parameter is an object handled by this converted library and pass the adress of the referenced structure
									t.="		If (IsObject(" ParamNames[A_Index+1] ") and (ComObjType(" ParamNames[A_Index+1] ")=""""))`r`n"
									t.="			ref" ParamNames[A_Index+1] ":=" ParamNames[A_Index+1] ".__Value`r`n"
									; If the parameter is a structure, pass it's address
									t.="		else`r`n"
									t.="			ref" ParamNames[A_Index+1] ":=&" ParamNames[A_Index+1] "`r`n"
								}								
							}
						}							
						t.="		res:=DllCall(this.__Vt(" FuncDescVar.oVft//A_PtrSize "), ""Ptr"", this.__Value"
						Loop, % FuncDescVar.cParams
						{
							param:=new ELEMDESC(FuncDescVar.lprgelemdescParam+(A_Index-1)*ELEMDESC.SizeOf())
								t.=", "
							If (param.tdesc.vt=12) ; VT_Variant requires different handling in 32/64Bit
							{ 
								If (A_PtrSize=4)
								{
									t.= """int64"", " ParamNames[A_Index+1] "VariantType, ""int64"", ref" ParamNames[A_Index+1]
								}									
								else
								{
									t.= """Ptr"", &ref" ParamNames[A_Index+1]
								}
							}
							else
							{
								to:=GetTypeObj(param, TI)
								t.=""""  GetAHKDllCallTypeFromVarObj(to) """, "
								If (ReturnIsHRESULT and HasRetValParam and (param.paramdesc.wParamFlags & 0x8))
									t.="out"
								else									
								If ((to[1].Type=26 and to[2].Type=29 and to[2].RefType=13) or (to[1].Type=29 and to[1].RefType=13))
									t.="ref" ParamNames[A_Index+1]
								else
								If ((to[1].Type=26 and to[2].Type=27) or (to[1].Type=27))
									t.="ComObjValue(ref" ParamNames[A_Index+1] ")"
								else
									t.=ParamNames[A_Index+1]
							}
						}
						t.=", """ VT2AHK(result) """)`r`n"
						
						If (ReturnIsHRESULT and HasRetValParam)
						{
							t.="		If (res<0 and res<>"""")`r`n			Throw Exception(""COM HResult: 0x"" Format(""{:x}"", res & 0xFFFFFFFF) "" from " Name " in " TI.GetDocumentation(-1).Name """)`r`n"
							for index, str in GetTypeObjPostProcessing(RetValTypeObj)
								t.="		" str "`r`n"
							t.="		return out`r`n"
						}
						else
						{
							t.="		return res`r`n"
						}							
						t.="	}`r`n`r`n"
					}
					else	; it's a property
					{
						If (!IsObject(PropertyStore[Name]))
						{
							PropertyStore[Name]:=Object()
							PropertyStore[Name].Get:=0
							PropertyStore[Name].Put:=0
							PropertyStore[Name].PutRef:=0
						}
						vtype:=""
						ParamNames:=TI.GetNames(FuncDescVar.memid,FuncDescVar.cParams+1)
						param:=new ELEMDESC(FuncDescVar.lprgelemdescParam)
						VType.=GetVarStr(param, TI)
						VObj:=GetTypeObj(param, TI)
						VName:=ParamNames[2]
						If (FuncDescVar.invkind=2)
						{
							PropertyStore[Name].Get:=FuncDescVar.oVft//A_PtrSize
							PropertyStore[Name].GetType:=VType
							PropertyStore[Name].GetTypeObj:=VObj
							PropertyStore[Name].GetName:=VName
							PropertyStore[Name].RequiresInitialization:=GetRequiresInitialization(param, TI)
						}
						else
						If (FuncDescVar.invkind=3)
						{
							PropertyStore[Name].Put:=FuncDescVar.oVft//A_PtrSize
							PropertyStore[Name].PutType:=VType
							PropertyStore[Name].PutTypeObj:=VObj
							PropertyStore[Name].PutName:=VName
						}
						else
						If (FuncDescVar.invkind=4)
						{
							PropertyStore[Name].PutRef:=FuncDescVar.oVft//A_PtrSize
							PropertyStore[Name].PutRefType:=VType
							PropertyStore[Name].PutRefTypeObj:=VObj
							PropertyStore[Name].PutRefName:=VName
						}
					}
				}
				TI.ReleaseFuncDesc(FuncDescVar.__Ptr)
				For Name, PropObj in PropertyStore
				{
					t.="	" "; " this.Resources.Property Name	 
					If (PropObj.Get<>0)
						t.="; " this.Resources.VTablePositon this.Resources.Get PropObj.Get "; output : " PropObj.GetType ": " PropObj.GetName 
					If (PropObj.Put<>0)
						t.="; " this.Resources.VTablePositon this.Resources.Put PropObj.Put "; input : " PropObj.PutType ": " PropObj.PutName 
					If (PropObj.PutRef<>0)
						t.="; " this.Resources.VTablePositon this.Resources.Put PropObj.PutRef "; input : " PropObj.PutRefType ": " PropObj.PutRefName 
					t.="`r`n"
					t.="	" Name "[]`r`n	{`r`n"
					If (PropObj.Get<>0)
					{
						t.="		get {`r`n"
						If (PropObj.RequiresInitialization<>"")
						{
							t.="			out:=new " PropObj.RequiresInitialization "()`r`n"
							t.="			If !DllCall(this.__Vt(" PropObj.Get "), ""Ptr"", this.__Value, """ GetAHKDllCallTypeFromVarObj(PropObj.GetTypeObj) """, out.__Value)`r`n"
						}
						else
							t.="			If !DllCall(this.__Vt(" PropObj.Get "), ""Ptr"", this.__Value, """ GetAHKDllCallTypeFromVarObj(PropObj.GetTypeObj) """,out)`r`n"
						t.="			{`r`n"
						for index, str in GetTypeObjPostProcessing(PropObj.GetTypeObj)
							t.="				" str "`r`n"
						t.="				return out`r`n"
						t.="			}`r`n		}`r`n"
					}
					If (PropObj.Put<>0)
					{
						t.="		set {`r`n			If !DllCall(this.__Vt(" PropObj.Put "), ""Ptr"", this.__Value, """ GetAHKDllCallTypeFromVarObj(PropObj.PutTypeObj) """, value)`r`n"
						t.="				return value`r`n		}`r`n"
					}
					If (PropObj.PutRef<>0)
					{
						t.="		set {`r`n			If !DllCall(this.__Vt(" PropObj.PutRef "), ""Ptr"", this.__Value, """ GetAHKDllCallTypeFromVarObj(PropObj.PutRefTypeObj) """, value)`r`n"
						t.="				return value`r`n		}`r`n"
					}
					t.="	}`r`n"
				}
			}
			TI.ReleaseTypeAttr(Attr.__Ptr)
			t.="}`r`n`r`n"
		}
		Progress, Off
		return t
	}
	
	MakeRecord(Type)	; Type "RECORD", "UNION"
	{
		t:=this.Resources.StartComment
		t.=this.Resources.StartComment
		t.="; " Type "`r`n"
		t.=this.Resources.StartComment
		t.=this.Resources.EndComment
		for index, obj in this.TLInfo[Type]
		{
			Progress, % (Index)/this.TLInfo[Type].MaxIndex()*100, % Index this.Resources.Of this.TLInfo.RECORD.MaxIndex(), % this.Resources.ProgressWindowText Type, % this.Resources.ProgressWindowCaption ": " this.TLInfo._TypeLibraryInfo.Name
			TI:=this.TypeLib.GetTypeInfo(obj.Index)
			t.=this.Resources.StartComment
			t.="; " TI.GetDocumentation(-1).Name	"`r`n"
			Attr:=TI.GetTypeAttr()
			t.="; " this.Resources.GUID Attr.guid "`r`n"
			t.=this.Resources.EndComment
			If (Attr.cVars)
			{
				If (this.Settings.OutPutV2)
					t.="class " TI.GetDocumentation(-1).Name "`r`n{`r`n	__New(byref p:=""empty"")`r`n	{`r`n		If (p=""empty"")`r`n		{`r`n			VarSetCapacity(p,this.__SizeOf(),0)`r`n		}`r`n		ObjRawSet(this, ""__Value"", &p)`r`n	}`r`n`r`n	__Get(VarName)`r`n	{`r`n		If (VarName=""__Value"")`r`n			return this__Value`r`n"
				else
					t.="class " TI.GetDocumentation(-1).Name "`r`n{`r`n	__New(byref p=""empty"")`r`n	{`r`n		If (p=""empty"")`r`n		{`r`n			VarSetCapacity(p,this.__SizeOf(),0)`r`n		}`r`n		ObjInsert(this, ""__Value"", &p)`r`n	}`r`n`r`n	__Get(VarName)`r`n	{`r`n		If (VarName=""__Value"")`r`n			return this.__Value`r`n"
				Loop, % Attr.cVars
				{
					VarDescVar:=TI.GetVarDesc(A_Index-1)
					Doc:=TI.GetDocumentation(VarDescVar.memid)
					If (VarDescVar.varkind=0) ; PerInstance
					{
						t.="		If (VarName=""" Doc.Name """)`r`n			return NumGet(this.__Value+" VarDescVar.lpvarvalue ", 0, """ VT2AHK(VarDescVar.elemdescVar.tdesc.vt) """) `; " this.Resources.Type GetVarStr(VarDescVar.elemdescVar, TI)
						If (Doc.DocString<>"")
							t.=": " Doc.DocString
						t.="`r`n"
					}
					TI.ReleaseVarDesc(VarDescVar.__Ptr)
				}
				t.="	}`r`n`r`n	__Set(VarName, byref Value)`r`n	{`r`n"
				Loop, % Attr.cVars
				{
					VarDescVar:=TI.GetVarDesc(A_Index-1)
					Doc:=TI.GetDocumentation(VarDescVar.memid)
					If (VarDescVar.varkind=0) ; PerInstance
					{
						t.="		If (VarName=""" Doc.Name """)`r`n			NumPut(Value, this.__Value+" VarDescVar.lpvarvalue ", 0, """ VT2AHK(VarDescVar.elemdescVar.tdesc.vt) """) `; " this.Resources.Type GetVarStr(VarDescVar.elemdescVar, TI)
						If (Doc.DocString<>"")
							t.=": " Doc.DocString
						t.="`r`n"
					}
					Size:=VarDescVar.lpvarvalue+VTSize(VarDescVar.elemdescVar.tdesc.vt)
					TI.ReleaseVarDesc(VarDescVar.__Ptr)
				}
				t.="		return Value`r`n	}`r`n`r`n	__SizeOf()`r`n	{`r`n		return " Size "`r`n	}`r`n}`r`n`r`n"
			}
			TI.ReleaseTypeAttr(Attr.__Ptr)
		}
		Progress, Off
		return t
	}
	
	MakeGeneric(Type)	; outputs the type library information as formatted comments
	{
		t:=this.Resources.StartComment
		t.=this.Resources.StartComment
		t.="; " Type "`r`n"
		t.=this.Resources.StartComment
		t.=this.Resources.EndComment
		for index, obj in this.TLInfo[type]
		{
			Progress, % (Index)/this.TLInfo[Type].MaxIndex()*100, % Index this.Resources.Of this.TLInfo[Type].MaxIndex(), % this.Resources.ProgressWindowText Type, % this.Resources.ProgressWindowCaption ": " this.TLInfo._TypeLibraryInfo.Name
			TI:=this.TypeLib.GetTypeInfo(obj.Index)
			t.=this.Resources.StartComment
			t.="; " TI.GetDocumentation(-1).Name	"`r`n"
			t.=this.Resources.EndComment
			Attr:=TI.GetTypeAttr()
			t.=this.Resources.StartBlockComment
			t.=this.Resources.GUID Attr.guid "`r`n`r`n"
			If (Attr.cVars)
			{
				Loop, % Attr.cVars
				{
					VarDescVar:=TI.GetVarDesc(A_Index-1)
					Doc:=TI.GetDocumentation(VarDescVar.memid)
					If (VarDescVar.varkind=0) ; PerInstance
					{
						t.=Doc.Name "; " this.Resources.Offset VarDescVar.lpvarvalue ", " this.Resources.Type GetVarStr(VarDescVar.elemdescVar, TI)
						If (Doc.DocString<>"")
							t.="		`; " Doc.DocString
						t.="`r`n"
					}
					else
					If (VarDescVar.varkind=2) ; Const
					{
						vs:=GetVarStr(VarDescVar.elemdescVar, TI)
						If (vs="Vt_Bstr")							
							t.=Doc.Name " := """  VarDescVar.oInst """  `; " this.Resources.Type vs ; "`r`n"
						else
							t.=Doc.Name " := "  VarDescVar.oInst "  `; " this.Resources.Type vs ; "`r`n"
						If (Doc.DocString<>"")
							 t.="  `; " Doc.DocString 
						t.="`r`n"
					}
					TI.ReleaseVarDesc(VarDescVar.__Ptr)
				}
			}
			If (Attr.cFuncs)
			{
				Loop, % Attr.cFuncs
				{
					FuncIndex:=A_Index-1
					FuncDescVar:=TI.GetFuncDesc(FuncIndex)
					Doc:=TI.GetDocumentation(FuncDescVar.memid)
					Name:=Doc.Name
					If Name not in QueryInterface,AddRef,Release,GetTypeInfoCount,GetTypeInfo,GetIDsOfNames,Invoke
					{
						t.=this.Resources.VTablePositon ": " FuncDescVar.oVft//A_PtrSize ":`r`n"
						t.=INVOKEKIND(FuncDescVar.invkind) " "
						t.=SubStr(VARENUM(FuncDescVar.elemdescFunc.tdesc.vt), 4) " "
						t.=Name "("
						ParamNames:=TI.GetNames(FuncDescVar.memid,FuncDescVar.cParams+1)
						Loop, % FuncDescVar.cParams
						{
							If (A_Index>1)
								t.=", "
							param:=new ELEMDESC(FuncDescVar.lprgelemdescParam+(A_Index-1)*ELEMDESC.SizeOf())
							for index, key in PARAMFLAG(param.paramdesc.wParamFlags)
								t.="[" SubStr(key, 11) "] "
							vstr:=GetVarStr(param, TI)
							If (InStr(vstr,"VT_")=1)
								t.=Substr(vstr, 4) ": "
							else
								t.=vstr ": "
							t.=ParamNames[A_Index+1]
							If (param.paramdesc.wParamFlags & 0x20) ; Hasdefault
							{
								t.=" = " param.paramdesc.pPARAMDescEx.varDefaultValue
							}
						}
						t.=")`r`n"
						If (Doc.DocString<>"")
							t.=Doc.DocString "`r`n"
						t.="`r`n"
					}
					TI.ReleaseFuncDesc(FuncDescVar.__Ptr)
				}
			}
			If (Attr.cImplTypes)
			{
				If (Attr.typekind=5)	; TKIND_COCLASS
				{
					Loop, % Attr.cImplTypes
					{
						ImplT:=TI.GetRefTypeOfImplType(A_Index-1)
						Name:=ImplT.GetDocumentation(-1).Name
						t.=this.Resources.Implements Name "`r`n"
						Attr:=ImplT.GetTypeAttr()
						t.="; " this.Resources.GUID Attr.guid "`r`n"
						t.=this.Resources.Flags
						for index, flag in IMPLTYPEFLAGS(TI.GetImplTypeFlags(A_Index-1))
						{
							If (index>1)
								t.=", "
							t.=flag
						}
					}
				}
				else
				{
					Loop, % Attr.cImplTypes
					{
						Name:=TI.GetRefTypeOfImplType(A_Index-1).GetDocumentation(-1).Name
						t.=this.Resources.Implements Name "`r`n"
					t.=this.Resources.Flags 
					for index, flag in IMPLTYPEFLAGS(TI.GetImplTypeFlags(A_Index-1))
					{
						If (index>1)
							t.=", "
						t.=flag
					}
					t.="`r`n"
					}
				}
			}
			If (Attr.typekind=6)	; TKIND_ALIAS
			{
				t.=this.Resources.Alias GetVarStrFromTD(Attr.tdescAlias, TI) "`r`n"
				for index, flag in IMPLTYPEFLAGS(TI.GetImplTypeFlags(A_Index-1))
				{
					If (index=1)
						t.=this.Resources.Flags 
					else
						t.=", "
					t.=flag
				}
				t.="`r`n"
			}
			TI.ReleaseTypeAttr(Attr.__Ptr)
			t.=this.Resources.EndBlockComment
		}
		Progress, Off
		return t
	}
}
