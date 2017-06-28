VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    Dim dTime As Date
    Dim mi As MailItem
    
    If TypeName(Item) = "MailItem" Then
        Set mi = Item
        
        If mi.Importance = olImportanceHigh Then
            mi.Subject = Trim(mi.Subject)
            mi.DeferredDeliveryTime = DateAdd("n", 1, Now)
            MsgBox Prompt:="Send in one minute.", buttons:=vbOKOnly
        Else
            dTime = getNextSendWindow()
            
            Select Case MsgBox(Prompt:="Send at: " & Format(dTime, "HH:MM"), buttons:=vbYesNoCancel)
                Case vbOK:      mi.DeferredDeliveryTime = dTime
                Case vbNo:      mi.DeferredDeliveryTime = DateAdd("n", 1, Now)
                                MsgBox Prompt:="Send in one minute.", buttons:=vbOKOnly
                Case vbCancel:  Cancel = True
            End Select
        End If
    End If
End Sub


Private Function getNextSendWindow() As Date
    Dim dNow As Date
    Dim lNext As Long
    
    Dim lVariance As Long
    
    Dim dSend As Date
    
    lVariance = Int((59 - 1 + 1) * Rnd + 1)
    dNow = Now - Date
    
    Select Case dNow:
        Case Is < TimeSerial(11, 0, 0):
            Debug.Print "Nooner"
            dSend = TimeSerial(11, lVariance, 0)
        Case Is < TimeSerial(17, 0, 0):
            Debug.Print "Afternooner"
            dSend = TimeSerial(16, lVariance, 0)
        Case Else:
            Debug.Print "Midnight rush!"
            dSend = TimeSerial(11, lVariance, 0) + DateSerial(0, 0, 1)
    End Select
    
    Debug.Print "Send at:", (Date + dSend)
    
    getNextSendWindow = (Date + dSend)
End Function


