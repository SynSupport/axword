VERSION 5.00
Begin VB.UserControl AXWordCTL 
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   ScaleHeight     =   3645
   ScaleWidth      =   3990
End
Attribute VB_Name = "AXWordCTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************
'This OCX wraps WINWORD to give DBL access to WINWORD.
'The implementation is not a true wrap because it was
'easier for me to understand what needed to be done using
'a straight forward approach.  Because VB has access
'to all of WINWORD, I used DBL to get data and VB to
'drive WINWORD.
'********************************************************

'Force explicit declaration of all variables
Option Explicit

'Default Values:
Const m_def_account = "NONE"
Const m_def_invs = ""
Const m_def_pmts = ""
Const m_def_pdue = ""
Const m_def_cdue = ""
Const m_def_tdue = ""
Const m_def_tdate = ""
Const m_def_status = ""
Const m_def_ad = ""
Const m_def_rep = ""
Const m_def_DOTFile = ""

'Variables:
Dim m_account As String
Dim m_invs As String
Dim m_pmts As String
Dim m_pdue As String
Dim m_cdue As String
Dim m_tdue As String
Dim m_tdate As String
Dim m_status As String
Dim m_ad As String
Dim m_rep As String
Dim m_DOTFile As String

Dim wapp As Word.Application

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_account = m_def_account
    m_invs = m_def_invs
    m_pmts = m_def_pmts
    m_pdue = m_def_pdue
    m_cdue = m_def_cdue
    m_tdue = m_def_tdue
    m_tdate = m_def_tdate
    m_status = m_def_status
    m_ad = m_def_ad
    m_rep = m_def_rep
    m_DOTFile = m_def_DOTFile
End Sub

'Get Account
Public Property Get account() As String
    account = m_account
End Property

'Set Account
Public Property Let account(ByVal New_account As String)
    m_account = New_account
End Property

'Set Unpaid Invoices
Public Function SetInvoice(ByVal Offset As Integer, ByVal New_invs As String) As Boolean
    Static size As Integer
    Static invs() As String
    Dim I As Integer
    
    On Error GoTo ErrorHandler
    
    If Offset > size Then
        size = Offset
        ReDim Preserve invs(size) As String
    End If
    
    invs(Offset) = New_invs + Chr(13)

    m_invs = ""
    For I = 1 To size Step 1
        m_invs = m_invs + invs(I)
    Next I
    
    'Return OK
    SetInvoice = False
    Exit Function
    
ErrorHandler:
    MsgBox ("Error in  SetInvoice(): " + Err.Source)
    
    'Return ERROR
    SetInvoice = True
    Exit Function

End Function

'Set Payments
Public Function SetPayment(ByVal Offset As Integer, ByVal New_pmts As String) As Boolean
    Static size As Integer
    Static pmts() As String
    Dim I As Integer
    
    On Error GoTo ErrorHandler
    
    If Offset > size Then
        size = Offset
        ReDim Preserve pmts(size) As String
    End If
    
    pmts(Offset) = New_pmts + Chr(13)

    m_pmts = ""
    For I = 1 To size Step 1
        m_pmts = m_pmts + pmts(I)
    Next I
    
    'Return OK
    SetPayment = False
    Exit Function
    
ErrorHandler:
    MsgBox ("Error in  SetPayment(): " + Err.Source)
    
    'Return ERROR
    SetPayment = True
    Exit Function

End Function

'Get Past Due
Public Property Get pdue() As String
    pdue = m_pdue
End Property

'Set Past Due
Public Property Let pdue(ByVal New_pdue As String)
    m_pdue = New_pdue
End Property

'Get Past Due
Public Property Get cdue() As String
    cdue = m_cdue
End Property

'Set Past Due
Public Property Let cdue(ByVal New_cdue As String)
    m_cdue = New_cdue
End Property

'Get Total Due
Public Property Get tdue() As String
    tdue = m_tdue
End Property

'Set Total Due
Public Property Let tdue(ByVal New_tdue As String)
    m_tdue = New_tdue
End Property

'Get To Date Total
Public Property Get tdate() As String
    tdate = m_tdate
End Property

'Set To Date Total
Public Property Let tdate(ByVal New_tdate As String)
    m_tdate = New_tdate
End Property

'Get Status
Public Property Get status() As String
    status = m_status
End Property

'Set Status
Public Property Let status(ByVal New_status As String)
    m_status = New_status
End Property

'Set Adddress
Public Function SetAddress(ByVal Offset As Integer, ByVal New_ad As String) As Boolean
    Static size As Integer
    Static ad() As String
    Dim I As Integer
    
    On Error GoTo ErrorHandler
    
    If Offset > size Then
        size = Offset
        ReDim Preserve ad(size) As String
    End If
    
    ad(Offset) = New_ad + Chr(13)

    m_ad = ""
    For I = 1 To size Step 1
        m_ad = m_ad + ad(I)
    Next I
    
    'Return OK
    SetAddress = False
    Exit Function
    
ErrorHandler:
    MsgBox ("Error in  SetInvoice(): " + Err.Source)
    
    'Return ERROR
    SetAddress = True
    Exit Function

End Function

'Get Rep
Public Property Get rep() As String
    rep = m_rep
End Property

'Set Rep
Public Property Let rep(ByVal New_rep As String)
    m_rep = New_rep
End Property

'Get DOTFile
Public Property Get DOTFile() As String
    DOTFile = m_DOTFile
End Property

'Set DOTFile
Public Property Let DOTFile(ByVal New_DOTFile As String)
    m_DOTFile = New_DOTFile
End Property

Private Sub UserControl_Initialize()
    On Error GoTo ErrorHandler
    
    'Startup WinWord
    Set wapp = CreateObject("Word.Application.8")
    
    'Do not fall into Error handler
    Exit Sub

ErrorHandler:
    MsgBox ("Error in UserControl_Initialize(): " + Err.Description)
    Exit Sub
    
End Sub

Public Function Generate() As Boolean
    'Generate Dunning Letter by searching for tokens
    'and replacing token with text.
    
    Dim wdoc As Document
    
    On Error GoTo ErrorHandler
    
    Set wdoc = wapp.Documents.Open(DOTFile, , True)
    
    Call replace("[ACCOUNT_ADDRESS]", m_ad)
    Call replace("[ACCOUNT_NUMBER]", m_account)
    Call replace("[TOTAL_DUE]", "$" + m_tdue)
    Call replace("[ACCOUNT_STATUS]", m_status)
    Call replace("[TO_DATE_AMOUNT]", "$" + m_tdate)
    Call replace("[PAST_DUE]", "$" + m_pdue)
    Call replace("[CURRENT_DUE]", "$" + m_cdue)
    Call replace("[CURRENT_INVOICES]", m_invs)
    Call replace("[ALL_PAYMENTS]", m_pmts)
    Call replace("[ACCOUNT_MANAGER]", m_rep)
    
    wapp.Visible = True
    
    'Return OK
    Generate = False
    Exit Function
    
ErrorHandler:
    MsgBox ("Error in Generate(): " + Err.Description)
    
    'Return ERROR
    Generate = True
    Exit Function
    
End Function

Private Sub replace(ByVal a_search As String, ByVal a_replace As String)
    'For information on find and replace see MSDN:
    '   Microsoft Ofice 97/Visual Basic Programmer's Guide
    '   Chapter 7: Microsoft Word Objects
    
    With wapp.Selection.Find
        .ClearFormatting
        .Text = a_search
        .Replacement.ClearFormatting
        .Replacement.Text = a_replace
        .MatchCase = True
        .Execute replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With

End Sub
