Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit




' This code was originally written by Dev Ashish.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of Dev Ashish


Private Declare Function apiGetComputerName Lib "kernel32" Alias _
"GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Function fOSMachineName() As String

'Returns the computername
    Dim lngLen As Long, lngX As Long
    Dim strCompName As String
    lngLen = 16
    strCompName = String$(lngLen, 0)
    lngX = apiGetComputerName(strCompName, lngLen)
    If lngX <> 0 Then
        fOSMachineName = Left$(strCompName, lngLen)
    Else
        fOSMachineName = ""
    End If
End Function


Public Function FormNamesToString(ByRef strForms As String, Optional ByVal chrDelimit As String = ";") As Integer
  ' Comments: Populate a string with names of all forms in the current database
  ' Sets    : strForms        String of form names to populate
  ' Params  : chrDelimit      Character to use as a delimiter, defaults to ; if not assigned
  ' Returns : Number of forms
  ' Source  : Total Visual SourceBook
  
  On Error GoTo PROC_ERR
  
  Dim intTotal As Integer
  Dim intForm As Integer
  
  intTotal = CurrentProject.AllForms.Count
  
  For intForm = 0 To intTotal - 1
    strForms = strForms & CurrentProject.AllForms(intForm).Name & chrDelimit
  Next intForm
  
  FormNamesToString = intTotal
  Debug.Print strForms
  
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modForms.FormNamesToString"
  Resume PROC_EXIT
End Function

Public Function FindUserName() As String
    ' Status:  In Devlopment
    ' Comments:
    ' Params  :
    ' Returns : String
    ' Created : 10/27/16 16:34 GB
    ' Modified:
    
    'TVCodeTools ErrorEnablerStart
    On Error GoTo PROC_ERR
    'TVCodeTools ErrorEnablerEnd

    FindUserName = fOSMachineName()

    'TVCodeTools ErrorHandlerStart
PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox Err.Description, vbCritical, "Module1.FindUserName"
    Resume PROC_EXIT
    'TVCodeTools ErrorHandlerEnd

End Function
