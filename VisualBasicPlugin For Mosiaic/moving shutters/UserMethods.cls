VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public gUserID As Long
Public gParentWnd As Long
Public MM As MMAppLib.UserCall


Public Function Startup(cmdLine As String) As Long


Set PubMM = MM

Load Form1
Form1.Show vbModal


End Function

Public Function DoCommand(cmdLine As String) As Long




End Function

Public Function Shutdown() As Long
End Function


