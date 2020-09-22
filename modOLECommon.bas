Attribute VB_Name = "modOLECommon"
Option Explicit

Public Declare Sub OleInitialize Lib "ole32.dll" (pvReserved As Any)
Public Declare Sub OleUninitialize Lib "ole32.dll" ()
Public Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As String, pclsid As modOLECommon.Guid) As Long
Public Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As String, lpiid As modOLECommon.Guid) As Long
Public Declare Function CoCreateInstance Lib "ole32.dll" (rclsid As modOLECommon.Guid, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As modOLECommon.Guid, ppv As Any) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function PutMem2 Lib "msvbvm60" (ByVal pWORDDst As Long, ByVal NewValue As Long) As Long
Private Declare Function PutMem4 Lib "msvbvm60" (ByVal pDWORDDst As Long, ByVal NewValue As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByVal pDWORDSrc As Long, ByVal pDWORDDst As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Const GMEM_FIXED As Long = &H0
Private Const asmPUSH_imm32 As Byte = &H68
Private Const asmRET_imm16 As Byte = &HC2
Private Const asmCALL_rel32 As Byte = &HE8

Public Type Guid
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Public Const unk_QueryInterface As Long = 0
Public Const unk_AddRef As Long = 1
Public Const unk_Release As Long = 2

Public Function CallInterface(ByVal pInterface As Long, ByVal Member As Long, ByVal ParamsCount As Long, Optional ByVal p1 As Long = 0, Optional ByVal p2 As Long = 0, Optional ByVal p3 As Long = 0, Optional ByVal p4 As Long = 0, Optional ByVal p5 As Long = 0, Optional ByVal p6 As Long = 0, Optional ByVal p7 As Long = 0, Optional ByVal p8 As Long = 0, Optional ByVal p9 As Long = 0, Optional ByVal p10 As Long = 0) As Long
  Dim i As Long, t As Long
  Dim hGlobal As Long, hGlobalOffset As Long
  
  If ParamsCount < 0 Then Err.Raise 5 'invalid call
  If pInterface = 0 Then Err.Raise 5
  
  '5 áàéò äëÿ çàïèõèâàíèÿ êàæäîãî ïàðàìåòðà â ñòåê
  '5 áàéò - PUSH this
  '5 áàéò - âûçîâ ìåìáåðà
  '3 áàéòà - ret 0x0010, âûïèõèâàÿ ïðè ýòîì è ïàðàìåòðû CallWindowProc
  '1 áàéò - âûðàâíèâàíèå, ïîñêîëüêó ïîñëåäíèé PutMem4 òðåáóåò 4 áàéòà.
  
  hGlobal = GlobalAlloc(GMEM_FIXED, 5 * ParamsCount + 5 + 5 + 3 + 1)
  If hGlobal = 0 Then Err.Raise 7 'insuff. memory
  hGlobalOffset = hGlobal
  
  If ParamsCount > 0 Then
    t = VarPtr(p1)
    For i = ParamsCount - 1 To 0 Step -1
      PutMem2 hGlobalOffset, asmPUSH_imm32
      hGlobalOffset = hGlobalOffset + 1
      GetMem4 t + i * 4, hGlobalOffset
      hGlobalOffset = hGlobalOffset + 4
    Next
  End If
  
  'Ïåðâûé ïàðàìåòð ëþáîãî èíòåðôåéñíîãî ìåòîäà - this. Äåëàåì...
  PutMem2 hGlobalOffset, asmPUSH_imm32
  hGlobalOffset = hGlobalOffset + 1
  PutMem4 hGlobalOffset, pInterface
  hGlobalOffset = hGlobalOffset + 4
  
  'Âûçîâ ìåìáåðà èíòåðôåéñà
  PutMem2 hGlobalOffset, asmCALL_rel32
  hGlobalOffset = hGlobalOffset + 1
  GetMem4 pInterface, VarPtr(t)     'äåðåôåðåíñ: íàõîäèì ïîëîæåíèå vTable
  GetMem4 t + Member * 4, VarPtr(t) 'ñìåùåíèå ïî vTable, ïîñëå ÷åãî äåðåôåðåíñ îíîãî
  PutMem4 hGlobalOffset, t - hGlobalOffset - 4
  hGlobalOffset = hGlobalOffset + 4

  'Èíòåðôåéñû stdcall. Ïîýòîìó íå áóäåì cdecl ó÷èòûâàòü.
    
  PutMem4 hGlobalOffset, &H10C2&        'ret 0x0010
  
  CallInterface = CallWindowProc(hGlobal, 0, 0, 0, 0)
  
  GlobalFree hGlobal
End Function
